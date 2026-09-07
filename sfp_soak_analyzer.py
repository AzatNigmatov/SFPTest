#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=======================================================================================
 QSFP28 100G Back-to-Back Soak Test Analyzer
 Анализатор логов приёмочного прогона б/у трансиверов QSFP28 на Juniper QFX5110-48S-4C
=======================================================================================

НАЗНАЧЕНИЕ
----------
Скрипт разбирает периодические (раз в 5 минут) снимки вывода команд

    show interfaces et-* extensive
    show interfaces diagnostics optics et-*

снятые одновременно с обоих концов линка в ходе soak-теста, и выдаёт:

  * Excel-книгу (.xlsx) с детализацией по модулям, DOM и временными рядами ошибок;
  * аналитический PDF-отчёт с графиками (прогрев, дрейф DOM, всплески FEC)
    и цветовой сводкой вердиктов.

МЕТОДОЛОГИЯ (важно)
-------------------
1. Счётчики Junos (FEC, carrier transitions, input errors) — КУМУЛЯТИВНЫЕ и живут
   на ПОРТУ, а не на модуле. Они не переносятся вместе с трансивером при ротации и
   могут быть обнулены (`clear interfaces statistics`) или сброшены при пересоздании
   интерфейса. Поэтому анализ строится на ПОИНТЕРВАЛЬНЫХ ДЕЛЬТАХ, а не на разнице
   «первое значение / последнее значение»: последняя даёт грубо неверный результат,
   если в середине прогона был сброс счётчика.

2. Момент физического монтажа/демонтажа модулей неизбежно порождает link down,
   всплеск carrier transitions и лавину FEC-ошибок. Эти артефакты НЕ являются
   свойством модуля. Скрипт выделяет «краевые окна» (--head-guard-min /
   --tail-guard-min) в начале и в конце каждого захвата и исключает их из расчёта
   метрик надёжности, вынося обнаруженные там события в отдельный раздел отчёта.

3. Интервал считается «грязным» (contaminated) и исключается из метрик, если:
      - на любом из его концов линк не в состоянии Up;
      - на любом из его концов есть Active defects, отличные от None;
      - произошёл сброс счётчика (отрицательная дельта);
      - интервал попал в краевое окно;
      - длительность интервала аномальна (пропуск в мониторинге).

4. Привязка к серийному номеру выполняется через внешнюю карту ротации
   (modules_map.csv), поскольку вывод `show interfaces ... extensive` серийных
   номеров НЕ содержит. Если карта не задана — используется псевдо-идентификатор
   сокета (тест/свитч/порт), а логика кросс-раундовой локализации отключается
   с явным предупреждением в отчёте.

ЛОГИКА РОТАЦИИ В 2 КРУГА
------------------------
  * модуль «грязный» в ДВУХ и более кругах с РАЗНЫМИ партнёрами  -> БРАК (подтверждён);
  * модуль «грязный» в одном круге и чистый в другом             -> ГОДЕН, виновата
                                                                    трасса/патч-корд/грязь;
  * модуль «грязный» и участвовал только в одном круге           -> ТРЕБУЕТ ПОВТОРА.

ЗАПУСК
------
    pip install pandas numpy matplotlib openpyxl reportlab
    python3 sfp_soak_analyzer.py --log-dir . --out-dir ./report

    # с картой ротации и профилем оптики:
    python3 sfp_soak_analyzer.py --map modules_map.csv --optic SR4

Автор: Network Automation / Data Analysis
=======================================================================================
"""

from __future__ import annotations

import argparse
import logging
import math
import os
import re
import sys
import textwrap
import traceback
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")  # headless-рендеринг, обязателен до импорта pyplot
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter

# ======================================================================================
#  1. КОНСТАНТЫ, ПОРОГИ И ПРОФИЛИ ОПТИКИ
# ======================================================================================

LOG = logging.getLogger("sfp_soak")

#: Номинальная линейная скорость 100GBASE-R с учётом FEC-оверхеда, бит/с.
#: Используется для пересчёта «FEC corrected errors/с» в оценочный pre-FEC BER.
LINE_RATE_BPS = 103.125e9

#: Целевой pre-FEC BER, с которым RS-FEC (Clause 91, FEC91) ещё уверенно справляется.
#: Порог взят как отраслевой ориентир 5e-5; наши пороги вердикта — доли от него.
FEC91_PREFEC_BER_BUDGET = 5.0e-5


@dataclass(frozen=True)
class OpticProfile:
    """Границы даташита для конкретного типа QSFP28-модуля (на линию/lane)."""

    name: str
    tx_dbm_min: float
    tx_dbm_max: float
    rx_dbm_min: float
    rx_dbm_max: float
    bias_ma_min: float
    bias_ma_max: float
    temp_c_min: float
    temp_c_max: float
    #: Типовой рабочий ток смещения — для детекта «уставшего» лазера.
    bias_ma_typ_max: float


#: 100GBASE-SR4 (850 нм, VCSEL, MMF). Ток смещения VCSEL — единицы мА.
PROFILE_SR4 = OpticProfile(
    name="100GBASE-SR4",
    tx_dbm_min=-8.4, tx_dbm_max=2.4,
    rx_dbm_min=-10.3, rx_dbm_max=2.4,
    bias_ma_min=1.0, bias_ma_max=13.0,
    temp_c_min=0.0, temp_c_max=70.0,
    bias_ma_typ_max=10.0,
)

#: 100GBASE-LR4 (1310 нм, EML, SMF). Ток смещения EML — десятки мА.
PROFILE_LR4 = OpticProfile(
    name="100GBASE-LR4",
    tx_dbm_min=-4.3, tx_dbm_max=4.5,
    rx_dbm_min=-10.6, rx_dbm_max=4.5,
    bias_ma_min=10.0, bias_ma_max=120.0,
    temp_c_min=0.0, temp_c_max=70.0,
    bias_ma_typ_max=90.0,
)

#: 100GBASE-CWDM4 / PSM4 — часто встречающийся «средний» вариант.
PROFILE_CWDM4 = OpticProfile(
    name="100G-CWDM4",
    tx_dbm_min=-6.5, tx_dbm_max=2.5,
    rx_dbm_min=-11.5, rx_dbm_max=2.5,
    bias_ma_min=10.0, bias_ma_max=120.0,
    temp_c_min=0.0, temp_c_max=70.0,
    bias_ma_typ_max=90.0,
)

OPTIC_PROFILES: Dict[str, OpticProfile] = {
    "SR4": PROFILE_SR4,
    "LR4": PROFILE_LR4,
    "CWDM4": PROFILE_CWDM4,
}


@dataclass
class Thresholds:
    """
    Пороги браковки. Вынесены в один объект, чтобы их можно было
    перекрыть из CLI без правки кода.
    """

    # --- FEC ------------------------------------------------------------------
    #: Оценочный pre-FEC BER, ниже которого линк считается «эталонно чистым».
    ber_pass_max: float = 1.0e-9
    #: Верхняя граница «рабочего, но без запаса» линка (10% бюджета FEC91).
    ber_warn_max: float = 5.0e-6
    #: Любой uncorrected > 0 в чистом интервале — безусловный брак.
    uncorrected_fail: int = 0
    #: Коэффициент вариации мгновенной скорости FEC corrected, выше которого
    #: поток ошибок считается «нестабильным» (всплесками) даже при низком BER.
    fec_rate_cv_warn: float = 2.0

    # --- Линк -----------------------------------------------------------------
    #: Допустимое число флапов за прогон вне краевых окон.
    link_flaps_fail: int = 0
    #: Допустимое число input/CRC/bit errors.
    input_errors_fail: int = 0

    # --- DOM ------------------------------------------------------------------
    #: Запас до границы даташита, ниже которого выставляется WARNING, дБ.
    dom_margin_warn_db: float = 1.0
    #: Минимальная длительность наблюдения, при которой наклон тренда вообще
    #: имеет смысл экстраполировать на сутки, ч. Наклон, снятый за 15 минут и
    #: умноженный на 96, — это шум, а не деградация: без этого порога любой
    #: обрезанный или короткий захват давал бы ложную отбраковку по дрейфу.
    min_hours_for_drift: float = 6.0
    #: Минимальное число снимков DOM для оценки перекоса между линиями.
    min_samples_for_dom_stats: int = 12
    #: Минимальная наработка в стабильном режиме, при которой модуль вообще
    #: может быть аттестован в ЗИП, ч. Отсутствие ошибок за 15 минут не
    #: доказывает ничего: методика soak-теста предполагает 24-48 часов.
    min_hours_for_pass: float = 12.0
    #: Полный дрейф TX/RX за прогон, приводящий к WARNING / FAIL, дБ.
    power_drift_warn_db: float = 1.0
    power_drift_fail_db: float = 2.0
    #: Разброс средней мощности между 4 линиями (перекос), дБ.
    lane_imbalance_warn_db: float = 2.0
    lane_imbalance_fail_db: float = 3.5
    #: Рост тока смещения, нормированный на прогон (мА / 24 ч).
    bias_drift_warn_ma_per_day: float = 0.30
    bias_drift_fail_ma_per_day: float = 0.80
    #: Рост тока смещения, не объяснимый прогревом (мА на градус), — «уставший» лазер.
    bias_per_degc_warn: float = 0.10


#: Вердикты, упорядоченные по «тяжести». Используется для агрегации.
VERDICT_PASS = "В ЗИП"
VERDICT_WARN = "Рабочий, не в ЗИП"
VERDICT_FAIL = "БРАК / Утилизация"
VERDICT_UNKNOWN = "Нет данных"

VERDICT_SEVERITY = {
    VERDICT_UNKNOWN: 0,
    VERDICT_PASS: 1,
    VERDICT_WARN: 2,
    VERDICT_FAIL: 3,
}

VERDICT_COLORS = {  # RGB hex без '#', для openpyxl / reportlab
    VERDICT_PASS: "C6EFCE",
    VERDICT_WARN: "FFEB9C",
    VERDICT_FAIL: "FFC7CE",
    VERDICT_UNKNOWN: "D9D9D9",
}
VERDICT_FONT_COLORS = {
    VERDICT_PASS: "006100",
    VERDICT_WARN: "9C6500",
    VERDICT_FAIL: "9C0006",
    VERDICT_UNKNOWN: "3F3F3F",
}


# ======================================================================================
#  2. РЕГУЛЯРНЫЕ ВЫРАЖЕНИЯ
# ======================================================================================

#: Разделитель снимков: "===== Tue Sep  1 16:31:20 +05 2026 ====="
RE_SNAPSHOT = re.compile(r"^={3,}\s*(?P<ts>.+?)\s*={3,}\s*$", re.MULTILINE)

#: Разбор самой метки времени. Не полагаемся на strptime с %Z: зона может быть
#: как символьной ("UTC", "MSK"), так и числовой ("+05", "+0500").
RE_TIMESTAMP = re.compile(
    r"^(?P<wday>[A-Za-z]{3})\s+"
    r"(?P<mon>[A-Za-z]{3})\s+"
    r"(?P<day>\d{1,2})\s+"
    r"(?P<hh>\d{1,2}):(?P<mm>\d{2}):(?P<ss>\d{2})\s+"
    r"(?P<tz>[A-Za-z]{2,5}|[+-]\d{2}:?\d{0,2})\s+"
    r"(?P<year>\d{4})\s*$"
)

MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}

#: Заголовок физического интерфейса.
RE_PHY_HEADER = re.compile(
    r"^Physical interface:\s*(?P<port>[A-Za-z]{2}-\d+/\d+/\d+)\s*,\s*"
    r"(?P<admin>\w+)\s*,\s*Physical link is\s+(?P<link>\w+)",
    re.MULTILINE,
)

#: Строка счётчиков входящих ошибок (все поля опциональны по факту наличия).
RE_CARRIER = re.compile(r"Carrier transitions:\s*(\d+)")
RE_IN_ERRORS = re.compile(r"Carrier transitions:\s*\d+\s*,\s*Errors:\s*(\d+)")
RE_IN_DROPS = re.compile(r"\bDrops:\s*(\d+)")
RE_HS_CRC = re.compile(r"HS link CRC errors:\s*(\d+)")
RE_MTU_ERR = re.compile(r"MTU errors:\s*(\d+)")
RE_RESOURCE = re.compile(r"Resource errors:\s*(\d+)")
RE_FIFO = re.compile(r"FIFO errors:\s*(\d+)")
RE_AGED = re.compile(r"Aged packets:\s*(\d+)")
RE_COLLISIONS = re.compile(r"Collisions:\s*(\d+)")

RE_DEFECTS = re.compile(r"Active defects\s*:\s*(?P<defects>.*?)\s*$", re.MULTILINE)
RE_BIT_ERRORS = re.compile(r"^\s*Bit errors\s+(\d+)\s*$", re.MULTILINE)
RE_ERRORED_BLOCKS = re.compile(r"^\s*Errored blocks\s+(\d+)\s*$", re.MULTILINE)
RE_FEC_MODE = re.compile(r"Ethernet FEC Mode\s*:\s*(\S+)")
RE_FEC_CORR = re.compile(r"^\s*FEC Corrected Errors\s+(\d+)\s*$", re.MULTILINE)
RE_FEC_UNCORR = re.compile(r"^\s*FEC Uncorrected Errors\s+(\d+)\s*$", re.MULTILINE)
RE_FEC_CORR_RATE = re.compile(r"^\s*FEC Corrected Errors Rate\s+(\d+)\s*$", re.MULTILINE)
RE_FEC_UNCORR_RATE = re.compile(r"^\s*FEC Uncorrected Errors Rate\s+(\d+)\s*$", re.MULTILINE)
#: "CRC/Align errors    0    0"  -> берём оба столбца (input / output).
RE_CRC_ALIGN = re.compile(r"^\s*CRC/Align errors\s+(\d+)\s+(\d+)\s*$", re.MULTILINE)

#: Заголовок секции оптической диагностики: "Diagnostics Optics interface Ethernet-0/0/48"
RE_OPTICS_HEADER = re.compile(
    r"^\s*Diagnostics Optics interface\s+(?P<ifname>[A-Za-z]+-\d+/\d+/\d+)\s*$",
    re.MULTILINE,
)

RE_MODULE_TEMP = re.compile(
    r"Module temperature\s*:\s*(?P<c>-?\d+(?:\.\d+)?)\s*degrees\s*C", re.IGNORECASE)
RE_MODULE_VOLT = re.compile(
    r"Module voltage\s*:\s*(?P<v>-?\d+(?:\.\d+)?)\s*V", re.IGNORECASE)

#: Маркер начала блока линии: "lane 1:" / "Lane 0:" / "Laser 1:"
RE_LANE_MARK = re.compile(r"^\s*(?:lane|laser|channel)\s*(?P<idx>\d+)\s*:\s*$",
                          re.IGNORECASE | re.MULTILINE)

RE_BIAS = re.compile(
    r"Laser bias current\s*:\s*(?P<val>-?\d+(?:\.\d+)?)\s*mA", re.IGNORECASE)

#: Мощность может быть представлена как "1.006 mW / 0.03 dBm", только "mW",
#: только "dBm", либо специальными значениями ("- Inf", "Off", "N/A").
RE_TX_POWER = re.compile(r"Laser output power\s*:\s*(?P<body>.+)$",
                         re.IGNORECASE | re.MULTILINE)
RE_RX_POWER = re.compile(r"Laser (?:receiver|rx) power\s*:\s*(?P<body>.+)$",
                         re.IGNORECASE | re.MULTILINE)

RE_MW_VALUE = re.compile(r"(?P<val>-?\d+(?:\.\d+)?)\s*mW", re.IGNORECASE)
RE_DBM_VALUE = re.compile(r"(?P<val>-?\s?(?:\d+(?:\.\d+)?|Inf))\s*dBm", re.IGNORECASE)

#: Имя файла лога -> сторона (A/B) и номер круга ротации.
#: Понимает: sw-a1.log, clean_sw-a1.log, sw-A2.log, soak2_sw-B.log, sw_b_test1.log ...
RE_FILENAME = re.compile(
    r"sw[-_]?(?P<side>[ABab])[-_]?(?P<test>\d+)?", re.IGNORECASE)
RE_FILENAME_TESTNUM = re.compile(r"(?:soak|test|round|krug|круг)[-_]?(?P<n>\d+)",
                                 re.IGNORECASE)

#: Строка приглашения PuTTY/консоли — используется, чтобы отбросить преамбулу.
RE_NOISE_LINE = re.compile(r"^(?:=~=~=|\S+@\S+[:%#>]|\s*$)")


# ======================================================================================
#  3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ======================================================================================

def safe_int(match: Optional[re.Match], group: int = 1) -> Optional[int]:
    """Достаёт целое из результата поиска, не падая на None/мусоре."""
    if match is None:
        return None
    try:
        return int(match.group(group))
    except (ValueError, IndexError, TypeError):
        return None


def safe_float(match: Optional[re.Match], group: str | int = 1) -> Optional[float]:
    """Достаёт float из результата поиска, не падая на None/мусоре."""
    if match is None:
        return None
    try:
        return float(match.group(group))
    except (ValueError, IndexError, TypeError):
        return None


def mw_to_dbm(mw: Optional[float]) -> Optional[float]:
    """Перевод мВт -> дБм. 0 и отрицательные значения трактуются как «нет света»."""
    if mw is None:
        return None
    if mw <= 0:
        return float("-inf")
    return 10.0 * math.log10(mw)


def dbm_to_mw(dbm: Optional[float]) -> Optional[float]:
    """Перевод дБм -> мВт."""
    if dbm is None or dbm == float("-inf"):
        return 0.0 if dbm == float("-inf") else None
    return 10.0 ** (dbm / 10.0)


def parse_power(body: str) -> Tuple[Optional[float], Optional[float]]:
    """
    Разбирает «хвост» строки мощности и возвращает кортеж (mW, dBm).

    Поддерживаются варианты:
        "1.006 mW / 0.03 dBm"   -> (1.006, 0.03)
        "0.893 mW"              -> (0.893, -0.49)   # dBm досчитывается
        "-2.35 dBm"             -> (0.582, -2.35)   # mW досчитывается
        "0.000 mW / - Inf dBm"  -> (0.0, -inf)
        "Off" / "N/A" / "-"     -> (None, None)
    """
    if body is None:
        return None, None
    text = body.strip()
    if not text or re.fullmatch(r"[-–—]|N/?A|Off|Unknown", text, re.IGNORECASE):
        return None, None

    mw: Optional[float] = None
    dbm: Optional[float] = None

    m_mw = RE_MW_VALUE.search(text)
    if m_mw:
        try:
            mw = float(m_mw.group("val"))
        except ValueError:
            mw = None

    m_dbm = RE_DBM_VALUE.search(text)
    if m_dbm:
        raw = m_dbm.group("val").replace(" ", "")
        if raw.lower().lstrip("-+").startswith("inf"):
            dbm = float("-inf") if raw.startswith("-") else float("inf")
        else:
            try:
                dbm = float(raw)
            except ValueError:
                dbm = None

    # Досчитываем недостающую половину пары.
    if dbm is None and mw is not None:
        dbm = mw_to_dbm(mw)
    if mw is None and dbm is not None:
        mw = dbm_to_mw(dbm)

    return mw, dbm


def parse_timestamp(raw: str) -> Tuple[Optional[datetime], Optional[str]]:
    """
    Разбирает метку времени вида "Tue Sep  1 16:31:20 +05 2026".

    Возвращает (naive datetime, метка временной зоны).

    ВАЖНО: время намеренно возвращается «наивным» (без tzinfo). В исследуемых
    логах стороны A и B помечены разными зонами (+05 и UTC), при этом стенные
    часы обоих коммутаторов идут синхронно — то есть метка зоны на одном из
    устройств выставлена некорректно. Приведение к абсолютному UTC развело бы
    сходственные события на 5 часов. Поэтому корреляция концов линка выполняется
    по стенному времени, а исходная метка зоны сохраняется в данных как атрибут.
    """
    if not raw:
        return None, None
    m = RE_TIMESTAMP.match(raw.strip())
    if not m:
        # Резервный путь: пробуем несколько распространённых форматов.
        for fmt in ("%a %b %d %H:%M:%S %Y", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
            try:
                return datetime.strptime(raw.strip(), fmt), None
            except ValueError:
                continue
        return None, None
    try:
        dt = datetime(
            year=int(m.group("year")),
            month=MONTHS[m.group("mon").title()],
            day=int(m.group("day")),
            hour=int(m.group("hh")),
            minute=int(m.group("mm")),
            second=int(m.group("ss")),
        )
    except (ValueError, KeyError):
        return None, m.group("tz")
    return dt, m.group("tz")


def normalize_port(ifname: str) -> str:
    """
    Приводит имя интерфейса к единому виду.

    В секции ошибок Junos печатает `et-0/0/48`, а в секции оптики —
    `Ethernet-0/0/48`. Сводим оба к каноническому `et-0/0/48`.
    """
    if not ifname:
        return ifname
    m = re.search(r"(\d+/\d+/\d+)", ifname)
    return f"et-{m.group(1)}" if m else ifname.strip()


def port_short(port: str) -> str:
    """`et-0/0/48` -> `48` (для компактных подписей на графиках)."""
    m = re.search(r"/(\d+)$", port or "")
    return m.group(1) if m else (port or "?")


def linear_slope(x: Sequence[float], y: Sequence[float]) -> Optional[float]:
    """
    МНК-наклон y(x). Возвращает None, если данных недостаточно
    или ряд вырожден (все x одинаковы, есть NaN/inf).
    """
    xs = np.asarray(x, dtype=float)
    ys = np.asarray(y, dtype=float)
    mask = np.isfinite(xs) & np.isfinite(ys)
    xs, ys = xs[mask], ys[mask]
    if xs.size < 3 or np.ptp(xs) == 0:
        return None
    try:
        return float(np.polyfit(xs, ys, 1)[0])
    except (np.linalg.LinAlgError, ValueError):
        return None


def fmt_num(value: Any, digits: int = 2, dash: str = "—") -> str:
    """Аккуратное форматирование числа для таблиц отчёта."""
    if value is None:
        return dash
    if isinstance(value, float):
        if math.isnan(value):
            return dash
        if math.isinf(value):
            return "-∞" if value < 0 else "+∞"
        if abs(value) >= 1000:
            return f"{value:,.0f}".replace(",", " ")
        return f"{value:.{digits}f}"
    if isinstance(value, (int, np.integer)):
        return f"{int(value):,}".replace(",", " ")
    return str(value)


def fmt_count(value: Any, dash: str = "—") -> str:
    """
    Форматирует счётчик как целое число.

    Счётчики ошибок и флапов хранятся во float (сумма дельт), но «0.00 флапов»
    в отчёте читается как измерение с точностью до сотых, чем не является.
    """
    if value is None:
        return dash
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(v):
        return dash
    if math.isinf(v):
        return "-\u221e" if v < 0 else "+\u221e"
    return f"{int(round(v)):,}".replace(",", "\u202f")


def fmt_ber(value: Optional[float]) -> str:
    """Научная запись для BER."""
    if value is None or (isinstance(value, float) and (math.isnan(value))):
        return "—"
    if value == 0:
        return "0"
    return f"{value:.2e}"


def human_duration(seconds: Optional[float]) -> str:
    """Секунды -> «25 ч 11 мин»."""
    if seconds is None or not math.isfinite(seconds):
        return "—"
    total = int(seconds)
    h, rem = divmod(total, 3600)
    m = rem // 60
    return f"{h} ч {m:02d} мин"


# ======================================================================================
#  4. ПАРСЕР ЛОГОВ
# ======================================================================================

@dataclass
class ParseIssue:
    """Одна зафиксированная проблема разбора — попадает в отдельный лист Excel."""

    file: str
    snapshot_index: Optional[int]
    timestamp: Optional[str]
    severity: str
    message: str


@dataclass
class CaptureMeta:
    """Метаданные одного файла-захвата."""

    path: Path
    side: str            # 'A' | 'B'
    test: int            # номер круга ротации
    switch: str          # 'sw-A' / 'sw-B'
    tz_label: Optional[str] = None
    snapshots: int = 0
    first_ts: Optional[datetime] = None
    last_ts: Optional[datetime] = None

    @property
    def label(self) -> str:
        return f"T{self.test}/{self.switch}"


class SoakLogParser:
    """
    Отказоустойчивый разборщик текстовых логов soak-теста.

    Принцип: ни одна ошибка в отдельно взятом снимке или порту не должна
    прерывать разбор файла. Все проблемы копятся в self.issues и выгружаются
    в отчёт, чтобы аналитик видел, какая часть данных недостоверна.
    """

    def __init__(self) -> None:
        self.issues: List[ParseIssue] = []
        self.captures: List[CaptureMeta] = []

    # ---------------------------------------------------------------- discovery ----

    @staticmethod
    def identify_capture(path: Path) -> Optional[Tuple[str, int]]:
        """
        Определяет (сторона, номер круга) по имени файла.

        `clean_sw-a1.log` -> ('A', 1);  `sw-B2.log` -> ('B', 2);
        `soak2_sw-B.log`  -> ('B', 2).
        """
        stem = path.stem
        m = RE_FILENAME.search(stem)
        if not m:
            return None
        side = m.group("side").upper()

        test: Optional[int] = None
        if m.group("test"):
            test = int(m.group("test"))
        else:
            m2 = RE_FILENAME_TESTNUM.search(stem)
            if m2:
                test = int(m2.group("n"))
        if test is None:
            test = 1  # разумный дефолт для одиночного прогона
        return side, test

    def discover(self, log_dir: Path, patterns: Sequence[str]) -> List[CaptureMeta]:
        """Находит файлы логов и раскладывает их по (круг, сторона)."""
        found: List[Path] = []
        for pattern in patterns:
            found.extend(sorted(log_dir.glob(pattern)))
        # Убираем дубликаты, сохраняя порядок.
        seen, unique = set(), []
        for p in found:
            if p.resolve() not in seen and p.is_file():
                seen.add(p.resolve())
                unique.append(p)

        captures: List[CaptureMeta] = []
        for path in unique:
            ident = self.identify_capture(path)
            if ident is None:
                self.issues.append(ParseIssue(
                    file=path.name, snapshot_index=None, timestamp=None,
                    severity="WARNING",
                    message="Не удалось определить сторону/круг по имени файла — файл пропущен.",
                ))
                LOG.warning("Пропускаю %s: не распознано имя файла", path.name)
                continue
            side, test = ident
            captures.append(CaptureMeta(path=path, side=side, test=test,
                                        switch=f"sw-{side}"))
        self.captures = captures
        return captures

    # ------------------------------------------------------------------ parsing ----

    def parse_capture(self, meta: CaptureMeta) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Разбирает один файл. Возвращает пару DataFrame: (порты, DOM).

        Стратегия: файл режется на снимки по разделителю `===== ... =====`,
        затем каждый снимок разбирается независимо в try/except.
        """
        try:
            text = meta.path.read_text(encoding="utf-8", errors="replace")
        except OSError as exc:
            self.issues.append(ParseIssue(
                file=meta.path.name, snapshot_index=None, timestamp=None,
                severity="ERROR", message=f"Не удалось прочитать файл: {exc}"))
            return pd.DataFrame(), pd.DataFrame()

        headers = [m.group("ts") for m in RE_SNAPSHOT.finditer(text)]
        # Первый элемент split — преамбула (баннер PuTTY, приглашение shell) — отбрасываем.
        chunks = RE_SNAPSHOT.split(text)
        # re.split с одной группой возвращает [pre, ts1, body1, ts2, body2, ...]
        bodies = chunks[2::2] if len(chunks) > 2 else []

        if not headers:
            self.issues.append(ParseIssue(
                file=meta.path.name, snapshot_index=None, timestamp=None,
                severity="ERROR",
                message="В файле не найдено ни одного снимка `===== <дата> =====`."))
            return pd.DataFrame(), pd.DataFrame()

        if len(bodies) != len(headers):
            self.issues.append(ParseIssue(
                file=meta.path.name, snapshot_index=None, timestamp=None,
                severity="WARNING",
                message=(f"Число заголовков ({len(headers)}) не совпало с числом тел "
                         f"снимков ({len(bodies)}); лишние отброшены.")))
            n = min(len(headers), len(bodies))
            headers, bodies = headers[:n], bodies[:n]

        port_rows: List[Dict[str, Any]] = []
        dom_rows: List[Dict[str, Any]] = []
        tz_labels: List[str] = []

        for idx, (raw_ts, body) in enumerate(zip(headers, bodies)):
            ts, tz = parse_timestamp(raw_ts)
            if tz:
                tz_labels.append(tz)
            if ts is None:
                self.issues.append(ParseIssue(
                    file=meta.path.name, snapshot_index=idx, timestamp=raw_ts,
                    severity="WARNING",
                    message="Не удалось разобрать метку времени — снимок пропущен."))
                continue
            try:
                p_rows, d_rows = self._parse_snapshot(meta, idx, ts, tz, body)
                port_rows.extend(p_rows)
                dom_rows.extend(d_rows)
            except Exception as exc:  # noqa: BLE001 — намеренно широкий перехват
                self.issues.append(ParseIssue(
                    file=meta.path.name, snapshot_index=idx, timestamp=raw_ts,
                    severity="ERROR",
                    message=f"Сбой разбора снимка: {exc.__class__.__name__}: {exc}"))
                LOG.debug("Трассировка:\n%s", traceback.format_exc())

        meta.snapshots = len(headers)
        meta.tz_label = max(set(tz_labels), key=tz_labels.count) if tz_labels else None

        df_ports = pd.DataFrame(port_rows)
        df_dom = pd.DataFrame(dom_rows)
        if not df_ports.empty:
            meta.first_ts = df_ports["ts"].min()
            meta.last_ts = df_ports["ts"].max()

        LOG.info("%-18s -> снимков: %3d | строк портов: %5d | строк DOM: %5d | зона: %s",
                 meta.path.name, meta.snapshots, len(df_ports), len(df_dom),
                 meta.tz_label or "?")
        return df_ports, df_dom

    # ---------------------------------------------------------- snapshot level ----

    def _parse_snapshot(
        self,
        meta: CaptureMeta,
        idx: int,
        ts: datetime,
        tz: Optional[str],
        body: str,
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Разбирает тело одного снимка: секцию интерфейсов + секцию оптики."""
        # Граница между двумя секциями — первое вхождение "Diagnostics Optics interface".
        first_optics = RE_OPTICS_HEADER.search(body)
        iface_region = body[: first_optics.start()] if first_optics else body
        optics_region = body[first_optics.start():] if first_optics else ""

        port_rows = self._parse_interfaces(meta, idx, ts, tz, iface_region)
        dom_rows = self._parse_optics(meta, idx, ts, tz, optics_region)
        return port_rows, dom_rows

    def _parse_interfaces(
        self, meta: CaptureMeta, idx: int, ts: datetime,
        tz: Optional[str], region: str,
    ) -> List[Dict[str, Any]]:
        """Разбирает блоки `Physical interface: ...` внутри одного снимка."""
        rows: List[Dict[str, Any]] = []
        matches = list(RE_PHY_HEADER.finditer(region))
        for i, m in enumerate(matches):
            start = m.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(region)
            block = region[start:end]
            port = normalize_port(m.group("port"))

            defects_m = RE_DEFECTS.search(block)
            defects = (defects_m.group("defects").strip() if defects_m else None)
            crc_m = RE_CRC_ALIGN.search(block)

            rows.append({
                "ts": ts,
                "tz": tz,
                "test": meta.test,
                "side": meta.side,
                "switch": meta.switch,
                "source_file": meta.path.name,
                "snapshot": idx,
                "port": port,
                "admin_state": m.group("admin"),
                "link_up": m.group("link").lower() == "up",
                "active_defects": defects,
                "defects_clean": (defects is None or defects.lower() == "none"),
                "carrier_transitions": safe_int(RE_CARRIER.search(block)),
                "input_errors": safe_int(RE_IN_ERRORS.search(block)),
                "input_drops": safe_int(RE_IN_DROPS.search(block)),
                "collisions": safe_int(RE_COLLISIONS.search(block)),
                "aged_packets": safe_int(RE_AGED.search(block)),
                "fifo_errors": safe_int(RE_FIFO.search(block)),
                "hs_link_crc_errors": safe_int(RE_HS_CRC.search(block)),
                "mtu_errors": safe_int(RE_MTU_ERR.search(block)),
                "resource_errors": safe_int(RE_RESOURCE.search(block)),
                "bit_errors": safe_int(RE_BIT_ERRORS.search(block)),
                "errored_blocks": safe_int(RE_ERRORED_BLOCKS.search(block)),
                "fec_mode": (RE_FEC_MODE.search(block).group(1)
                             if RE_FEC_MODE.search(block) else None),
                "fec_corrected": safe_int(RE_FEC_CORR.search(block)),
                "fec_uncorrected": safe_int(RE_FEC_UNCORR.search(block)),
                "fec_corrected_rate_dev": safe_int(RE_FEC_CORR_RATE.search(block)),
                "fec_uncorrected_rate_dev": safe_int(RE_FEC_UNCORR_RATE.search(block)),
                "crc_align_in": safe_int(crc_m, 1),
                "crc_align_out": safe_int(crc_m, 2),
            })

        if not matches:
            self.issues.append(ParseIssue(
                file=meta.path.name, snapshot_index=idx, timestamp=str(ts),
                severity="WARNING",
                message="В снимке нет ни одного блока `Physical interface:`."))
        return rows

    def _parse_optics(
        self, meta: CaptureMeta, idx: int, ts: datetime,
        tz: Optional[str], region: str,
    ) -> List[Dict[str, Any]]:
        """
        Разбирает блоки `Diagnostics Optics interface ...`.

        Внутри блока модульные параметры (температура, напряжение) идут до
        первого маркера `lane N:`, а далее по каждой линии — bias/TX/RX.
        Поддерживается и «безлинейный» вывод (одноканальные модули): тогда
        параметры относятся к условной линии 1.
        """
        rows: List[Dict[str, Any]] = []
        if not region:
            return rows

        heads = list(RE_OPTICS_HEADER.finditer(region))
        for i, hm in enumerate(heads):
            start = hm.start()
            end = heads[i + 1].start() if i + 1 < len(heads) else len(region)
            block = region[start:end]
            port = normalize_port(hm.group("ifname"))

            temp_c = safe_float(RE_MODULE_TEMP.search(block), "c")
            volt_v = safe_float(RE_MODULE_VOLT.search(block), "v")

            lane_marks = list(RE_LANE_MARK.finditer(block))
            if lane_marks:
                segments = []
                for j, lm in enumerate(lane_marks):
                    s = lm.end()
                    e = lane_marks[j + 1].start() if j + 1 < len(lane_marks) else len(block)
                    segments.append((int(lm.group("idx")), block[s:e]))
            else:
                # Одноканальный / нестандартный вывод — вся полезная часть как lane 1.
                body_start = hm.end()
                segments = [(1, block[body_start - start:])]

            for lane_idx, seg in segments:
                tx_m = RE_TX_POWER.search(seg)
                rx_m = RE_RX_POWER.search(seg)
                tx_mw, tx_dbm = parse_power(tx_m.group("body") if tx_m else None)
                rx_mw, rx_dbm = parse_power(rx_m.group("body") if rx_m else None)
                rows.append({
                    "ts": ts,
                    "tz": tz,
                    "test": meta.test,
                    "side": meta.side,
                    "switch": meta.switch,
                    "source_file": meta.path.name,
                    "snapshot": idx,
                    "port": port,
                    "lane": lane_idx,
                    "temp_c": temp_c,
                    "voltage_v": volt_v,
                    "bias_ma": safe_float(RE_BIAS.search(seg), "val"),
                    "tx_mw": tx_mw,
                    "tx_dbm": tx_dbm,
                    "rx_mw": rx_mw,
                    "rx_dbm": rx_dbm,
                })
        return rows

    # ----------------------------------------------------------------- driver ----

    def parse_all(self, captures: Sequence[CaptureMeta]) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Разбирает все найденные захваты и склеивает результат."""
        ports_frames, dom_frames = [], []
        for meta in captures:
            dfp, dfd = self.parse_capture(meta)
            if not dfp.empty:
                ports_frames.append(dfp)
            if not dfd.empty:
                dom_frames.append(dfd)

        df_ports = (pd.concat(ports_frames, ignore_index=True)
                    if ports_frames else pd.DataFrame())
        df_dom = (pd.concat(dom_frames, ignore_index=True)
                  if dom_frames else pd.DataFrame())

        for df in (df_ports, df_dom):
            if not df.empty:
                df.sort_values(["test", "switch", "port", "ts"], inplace=True,
                               kind="mergesort")
                df.reset_index(drop=True, inplace=True)
        return df_ports, df_dom


# ======================================================================================
#  5. КАРТА РОТАЦИИ (привязка сокет -> серийный номер)
# ======================================================================================

SOCKET_COLS = ["test", "switch", "port"]


def socket_id(test: Any, switch: Any, port: Any) -> str:
    """Канонический идентификатор «посадочного места»: `T1|sw-A|et-0/0/48`."""
    return f"T{test}|{switch}|{port}"


class ModuleMap:
    """
    Карта соответствия «сокет (круг/свитч/порт) -> серийный номер модуля».

    Вывод `show interfaces ... extensive` серийных номеров не содержит, поэтому
    привязка задаётся оператором извне. При отсутствии карты используется
    псевдо-серийник, равный идентификатору сокета, и кросс-раундовая логика
    ротации отключается.
    """

    def __init__(self, mapping: Optional[Dict[str, str]] = None,
                 partners: Optional[Dict[str, str]] = None) -> None:
        self.mapping: Dict[str, str] = mapping or {}
        #: Необязательная явная карта линков: сокет -> сокет-партнёр.
        self.partners: Dict[str, str] = partners or {}
        self.loaded_from: Optional[str] = None

    # ------------------------------------------------------------------ загрузка --

    @classmethod
    def load(cls, path: Optional[Path]) -> "ModuleMap":
        """
        Читает CSV с колонками: test, switch, port, serial [, partner_switch, partner_port].

        Отсутствие файла — не ошибка: возвращается пустая карта.
        """
        if path is None:
            return cls()
        if not path.exists():
            LOG.warning("Карта модулей %s не найдена — привязка к S/N недоступна.", path)
            return cls()
        try:
            df = pd.read_csv(path, dtype=str, comment="#").fillna("")
        except Exception as exc:  # noqa: BLE001
            LOG.error("Не удалось прочитать карту модулей %s: %s", path, exc)
            return cls()

        df.columns = [c.strip().lower() for c in df.columns]
        required = {"test", "switch", "port", "serial"}
        if not required.issubset(df.columns):
            LOG.error("В карте модулей нет обязательных колонок %s — карта игнорируется.",
                      sorted(required - set(df.columns)))
            return cls()

        mapping, partners = {}, {}
        for _, row in df.iterrows():
            try:
                test = int(str(row["test"]).strip())
            except ValueError:
                continue
            switch = str(row["switch"]).strip()
            port = normalize_port(str(row["port"]).strip())
            serial = str(row["serial"]).strip()
            if not serial:
                continue
            sid = socket_id(test, switch, port)
            mapping[sid] = serial
            if "partner_switch" in df.columns and "partner_port" in df.columns:
                psw = str(row.get("partner_switch", "")).strip()
                ppt = normalize_port(str(row.get("partner_port", "")).strip())
                if psw and ppt:
                    partners[sid] = socket_id(test, psw, ppt)

        obj = cls(mapping, partners)
        obj.loaded_from = str(path)
        LOG.info("Карта модулей загружена: %d записей из %s", len(mapping), path)
        return obj

    # ------------------------------------------------------------------ шаблон ---

    @staticmethod
    def write_template(path: Path, sockets: Iterable[Tuple[int, str, str]]) -> None:
        """Создаёт CSV-шаблон карты со всеми обнаруженными сокетами и пустыми S/N."""
        lines = [
            "# Карта ротации модулей QSFP28.",
            "# Заполните колонку serial реальными серийными номерами трансиверов.",
            "# Один и тот же serial, встреченный в разных test, включает логику",
            "# кросс-раундовой локализации (модуль vs патч-корд).",
            "# Колонки partner_switch/partner_port необязательны: если не заданы,",
            "# партнёр по линку определяется автоматически по совпадению номера порта.",
            "test,switch,port,serial,partner_switch,partner_port",
        ]
        for test, switch, port in sorted(sockets):
            lines.append(f"{test},{switch},{port},,,")
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        LOG.warning("Создан шаблон карты модулей: %s — заполните колонку 'serial'.", path)

    # ------------------------------------------------------------------ запросы --

    def serial_for(self, test: Any, switch: Any, port: Any) -> str:
        """Возвращает S/N либо псевдо-идентификатор сокета."""
        sid = socket_id(test, switch, port)
        return self.mapping.get(sid, sid)

    def is_real(self, test: Any, switch: Any, port: Any) -> bool:
        """True, если для сокета задан настоящий серийный номер."""
        return socket_id(test, switch, port) in self.mapping

    @property
    def active(self) -> bool:
        return bool(self.mapping)


def resolve_partner(
    df_sockets: pd.DataFrame, test: Any, switch: Any, port: Any,
    module_map: ModuleMap,
) -> Optional[Tuple[Any, Any, Any]]:
    """
    Ищет сокет на противоположном конце линка.

    Приоритет: явное указание в карте -> совпадение номера порта на другой стороне
    в том же круге. Возвращает (test, switch, port) партнёра или None.
    """
    sid = socket_id(test, switch, port)
    if sid in module_map.partners:
        target = module_map.partners[sid]
        row = df_sockets[df_sockets["socket_id"] == target]
        if not row.empty:
            r = row.iloc[0]
            return r["test"], r["switch"], r["port"]

    same_test = df_sockets[(df_sockets["test"] == test)
                           & (df_sockets["switch"] != switch)
                           & (df_sockets["port"] == port)]
    if len(same_test) == 1:
        r = same_test.iloc[0]
        return r["test"], r["switch"], r["port"]
    return None


# ======================================================================================
#  6. ДВИЖОК МЕТРИК: ПОИНТЕРВАЛЬНЫЕ ДЕЛЬТЫ, КРАЕВЫЕ ОКНА, СБРОСЫ СЧЁТЧИКОВ
# ======================================================================================

#: Кумулятивные счётчики, для которых считаются дельты.
COUNTER_COLS = [
    "carrier_transitions", "input_errors", "input_drops", "collisions",
    "aged_packets", "fifo_errors", "hs_link_crc_errors", "mtu_errors",
    "resource_errors", "bit_errors", "errored_blocks",
    "fec_corrected", "fec_uncorrected", "crc_align_in", "crc_align_out",
]

#: Счётчики, любой ненулевой прирост которых является браковочным признаком.
HARD_ERROR_COLS = [
    "input_errors", "hs_link_crc_errors", "mtu_errors", "resource_errors",
    "fifo_errors", "bit_errors", "errored_blocks", "crc_align_in", "crc_align_out",
]


def mark_edge_windows(
    df_ports: pd.DataFrame, head_guard_min: float, tail_guard_min: float,
) -> pd.DataFrame:
    """
    Помечает снимки, попавшие в «краевые окна» захвата.

    Мотивация: установка и извлечение трансиверов оператором в начале и в конце
    прогона неизбежно генерируют link down, скачки carrier transitions, обвал
    принимаемой мощности до уровня «нет света» и лавину FEC-ошибок. Это артефакты
    процедуры, а не свойства модуля. Окна считаются отдельно для каждого захвата
    (файла), поскольку стороны A и B стартуют и финишируют не строго синхронно.

    Функция применяется и к счётчикам ошибок, и к телеметрии DOM: без этого
    единственный снимок с извлечённым модулем даёт RX порядка -23 дБм и ложную
    отбраковку по выходу за границы даташита.

    Добавляет булевы колонки `in_head_window`, `in_tail_window`, `edge`.
    """
    df = df_ports.copy()
    df["in_head_window"] = False
    df["in_tail_window"] = False

    if df.empty:
        df["edge"] = False
        return df

    for (test, switch), grp in df.groupby(["test", "switch"], sort=False):
        t0, t1 = grp["ts"].min(), grp["ts"].max()
        head_edge = t0 + timedelta(minutes=head_guard_min)
        tail_edge = t1 - timedelta(minutes=tail_guard_min)
        idx = grp.index
        if head_guard_min > 0:
            df.loc[idx[grp["ts"] < head_edge], "in_head_window"] = True
        if tail_guard_min > 0:
            df.loc[idx[grp["ts"] > tail_edge], "in_tail_window"] = True

    df["edge"] = df["in_head_window"] | df["in_tail_window"]
    return df


def classify_interval(row: pd.Series) -> str:
    """
    Классифицирует интервал по ДОКАЗАТЕЛЬСТВАМ из самого лога, а не по догадкам
    о том, что делал оператор.

    Категории (в порядке приоритета):
      LINK_DOWN — на одном из концов линк не Up или есть Active defects:
                  ошибки набраны при потере сигнала;
      RESET     — счётчик уменьшился: интерфейс пересоздан или выполнен
                  `clear interfaces statistics`; значение после сброса — это
                  накопление «с нуля», а не продолжение ряда;
      FLAP      — вырос carrier transitions при поднятом линке: линк передёрнули
                  (пересадка модуля, переобучение), ошибки относятся к bring-up;
      GAP       — аномально длинный интервал: мониторинг не работал, данные неполны;
      STABLE    — линк Up, дефектов нет, флапов нет, сбросов нет.

    Ошибки в STABLE — это свойство тракта. Ошибки в остальных категориях
    требуют инженерного решения, но ЗАМАЛЧИВАТЬ их нельзя ни в каком случае.
    """
    if row.get("link_disturbed"):
        return "LINK_DOWN"
    if row.get("counter_reset"):
        return "RESET"
    if float(row.get("delta_carrier_transitions") or 0) > 0:
        return "FLAP"
    if row.get("time_gap"):
        return "GAP"
    return "STABLE"


#: Человекочитаемые названия категорий для отчётов.
ATTRIBUTION_LABELS = {
    "STABLE": "стабильный прогон",
    "FLAP": "передёргивание линка / bring-up",
    "RESET": "после сброса счётчика",
    "LINK_DOWN": "при потере линка",
    "GAP": "пропуск мониторинга",
}


def compute_intervals(df_ports: pd.DataFrame) -> pd.DataFrame:
    """
    Превращает ряд кумулятивных счётчиков в ряд поинтервальных дельт.

    Для каждого сокета (test, switch, port) строки сортируются по времени, затем:
      * `delta_<counter>` = разница со следующим снимком;
      * `counter_reset`   = True, если любая дельта отрицательна (обнуление
                            статистики или пересоздание интерфейса);
      * `link_disturbed`  = True, если на любом конце интервала линк не Up
                            либо присутствуют Active defects;
      * `time_gap`        = True, если длительность интервала аномально велика;
      * `attribution`     = категория из classify_interval();
      * `stable`          = интервал прошёл в штатном режиме.

    ВАЖНО: интервалы НЕ выбрасываются. Ни одна ошибка не исчезает из учёта —
    меняется только то, в какую категорию она попадёт. Скорости и BER считаются
    по стабильным интервалам (иначе они не имеют смысла), а жёсткие счётчики
    ошибок — по всему захвату целиком.
    """
    if df_ports.empty:
        return pd.DataFrame()

    frames: List[pd.DataFrame] = []
    for (test, switch, port), grp in df_ports.groupby(SOCKET_COLS, sort=False):
        g = grp.sort_values("ts", kind="mergesort").reset_index(drop=True)
        if len(g) < 2:
            continue

        cur, nxt = g.iloc[:-1].reset_index(drop=True), g.iloc[1:].reset_index(drop=True)
        iv = pd.DataFrame({
            "test": test,
            "switch": switch,
            "port": port,
            "socket_id": socket_id(test, switch, port),
            "ts_start": cur["ts"],
            "ts_end": nxt["ts"],
        })
        iv["duration_s"] = (iv["ts_end"] - iv["ts_start"]).dt.total_seconds()

        reset_flag = pd.Series(False, index=iv.index)
        for col in COUNTER_COLS:
            if col not in g.columns:
                iv[f"delta_{col}"] = np.nan
                continue
            a = pd.to_numeric(cur[col], errors="coerce")
            b = pd.to_numeric(nxt[col], errors="coerce")
            d = b - a
            reset_flag |= d.fillna(0) < 0
            # Отрицательная дельта означает сброс: значение ПОСЛЕ сброса — это
            # накопление с нуля, и именно оно является приростом за интервал.
            iv[f"delta_{col}"] = d.where(d >= 0, b)

        link_bad = (~cur["link_up"].fillna(False).astype(bool)
                    | ~nxt["link_up"].fillna(False).astype(bool))
        defects_bad = (~cur["defects_clean"].fillna(True).astype(bool)
                       | ~nxt["defects_clean"].fillna(True).astype(bool))
        iv["link_disturbed"] = (link_bad | defects_bad).values
        iv["counter_reset"] = reset_flag.values
        iv["edge"] = (cur["edge"].fillna(False).astype(bool).values
                      | nxt["edge"].fillna(False).astype(bool).values)

        med = iv["duration_s"].median()
        iv["time_gap"] = (iv["duration_s"] > max(3.0 * med, med + 300.0)
                          if med and math.isfinite(med) else False)

        iv["attribution"] = iv.apply(classify_interval, axis=1)
        iv["stable"] = iv["attribution"] == "STABLE"
        # Совместимость с прежним полем: пригодность для расчёта СКОРОСТЕЙ.
        iv["usable"] = iv["stable"]
        frames.append(iv)

    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    out.sort_values(["test", "switch", "port", "ts_start"], inplace=True,
                    kind="mergesort")
    out.reset_index(drop=True, inplace=True)
    return out


def build_error_events(df_iv: pd.DataFrame) -> pd.DataFrame:
    """
    Реестр КАЖДОГО события ошибки: интервал, в котором вырос хоть один жёсткий
    счётчик или счётчик carrier transitions.

    Это основной инструмент разбора: инженер видит, что именно произошло, когда,
    и в каком состоянии был линк в этот момент, — и может сам решить, списывать
    ли событие на пересадку модуля. Раньше эта информация терялась.
    """
    if df_iv.empty:
        return pd.DataFrame()

    watch = ["fec_uncorrected", "carrier_transitions", "fec_corrected"] + HARD_ERROR_COLS
    cols = [f"delta_{c}" for c in watch if f"delta_{c}" in df_iv.columns]
    if not cols:
        return pd.DataFrame()

    mask = pd.Series(False, index=df_iv.index)
    for c in cols:
        if c == "delta_fec_corrected":
            continue  # исправленные считаем отдельно, они есть почти всегда
        mask |= pd.to_numeric(df_iv[c], errors="coerce").fillna(0) > 0

    ev = df_iv[mask].copy()
    if ev.empty:
        return pd.DataFrame()

    out = pd.DataFrame({
        "test": ev["test"],
        "switch": ev["switch"],
        "port": ev["port"],
        "socket_id": ev["socket_id"],
        "ts_start": ev["ts_start"],
        "ts_end": ev["ts_end"],
        "attribution": ev["attribution"].map(ATTRIBUTION_LABELS).fillna(ev["attribution"]),
        "attribution_code": ev["attribution"],
        "fec_uncorrected": pd.to_numeric(ev.get("delta_fec_uncorrected"),
                                         errors="coerce").fillna(0),
        "fec_corrected": pd.to_numeric(ev.get("delta_fec_corrected"),
                                       errors="coerce").fillna(0),
        "carrier_transitions": pd.to_numeric(ev.get("delta_carrier_transitions"),
                                             errors="coerce").fillna(0),
        "counter_reset": ev["counter_reset"],
        "link_disturbed": ev["link_disturbed"],
    })
    for c in HARD_ERROR_COLS:
        key = f"delta_{c}"
        if key in ev.columns:
            out[c] = pd.to_numeric(ev[key], errors="coerce").fillna(0)
    out.sort_values(["test", "switch", "port", "ts_start"], inplace=True,
                    kind="mergesort")
    out.reset_index(drop=True, inplace=True)
    return out


def summarize_error_metrics(
    df_ports: pd.DataFrame, df_iv: pd.DataFrame,
) -> pd.DataFrame:
    """
    Агрегирует метрики надёжности по каждому сокету.

    Принцип разделения, исправляющий ключевой методический дефект:

      * ЖЁСТКИЕ счётчики (FEC uncorrected, флапы, input/CRC/bit errors)
        суммируются по ВСЕМУ захвату. Ни одно окно, ни один фильтр не может
        обнулить их: это первичные признаки браковки.
      * СКОРОСТИ и оценочный BER считаются только по стабильным интервалам —
        делить лавину ошибок момента пересадки модуля на время бессмысленно.
      * Дополнительно тот же жёсткий счётчик раскладывается по категориям
        (стабильный прогон / флап / сброс / потеря линка), чтобы инженер видел
        обстоятельства, а не только итог.
    """
    if df_ports.empty:
        return pd.DataFrame()

    rows: List[Dict[str, Any]] = []
    iv_by_socket = (df_iv.groupby("socket_id") if not df_iv.empty else None)

    for (test, switch, port), grp in df_ports.groupby(SOCKET_COLS, sort=False):
        sid = socket_id(test, switch, port)
        g = grp.sort_values("ts", kind="mergesort")

        try:
            iv = iv_by_socket.get_group(sid) if iv_by_socket is not None else pd.DataFrame()
        except KeyError:
            iv = pd.DataFrame()

        stable = iv[iv["stable"]] if not iv.empty else pd.DataFrame()

        def dsum(col: str, frame: pd.DataFrame) -> float:
            key = f"delta_{col}"
            if frame.empty or key not in frame.columns:
                return 0.0
            return float(pd.to_numeric(frame[key], errors="coerce").fillna(0)
                         .clip(lower=0).sum())

        def by_attr(col: str) -> Dict[str, float]:
            if iv.empty:
                return {k: 0.0 for k in ATTRIBUTION_LABELS}
            return {k: dsum(col, iv[iv["attribution"] == k]) for k in ATTRIBUTION_LABELS}

        # --- жёсткие счётчики: ВЕСЬ захват ------------------------------------
        fec_uncorr_total = dsum("fec_uncorrected", iv)
        fec_corr_total = dsum("fec_corrected", iv)
        ct_total = dsum("carrier_transitions", iv)
        unc_attr = by_attr("fec_uncorrected")
        ct_attr = by_attr("carrier_transitions")

        # --- скорости: только стабильные интервалы ----------------------------
        stable_seconds = float(stable["duration_s"].sum()) if not stable.empty else 0.0
        fec_corr_stable = dsum("fec_corrected", stable)
        corr_rate = (fec_corr_stable / stable_seconds if stable_seconds > 0
                     else float("nan"))
        ber_pre_fec = corr_rate / LINE_RATE_BPS if stable_seconds > 0 else float("nan")

        if not stable.empty and "delta_fec_corrected" in stable.columns:
            inst = (pd.to_numeric(stable["delta_fec_corrected"], errors="coerce")
                    .clip(lower=0) / stable["duration_s"].replace(0, np.nan))
            inst = inst.replace([np.inf, -np.inf], np.nan).dropna()
        else:
            inst = pd.Series(dtype=float)
        rate_max = float(inst.max()) if not inst.empty else float("nan")
        rate_mean = float(inst.mean()) if not inst.empty else float("nan")
        rate_std = float(inst.std(ddof=0)) if len(inst) > 1 else 0.0
        rate_cv = (rate_std / rate_mean) if (rate_mean and rate_mean > 0) else 0.0

        # --- флапы: весь захват -----------------------------------------------
        link_flaps = int(math.ceil(ct_total / 2.0)) if ct_total > 0 else 0
        down_samples = int((~g["link_up"].fillna(True).astype(bool)).sum())
        defect_samples = int((~g["defects_clean"].fillna(True).astype(bool)).sum())

        hard_errors = {c: dsum(c, iv) for c in HARD_ERROR_COLS}

        # --- целостность счётчиков --------------------------------------------
        first_unc = pd.to_numeric(g["fec_uncorrected"], errors="coerce").dropna()
        first_corr = pd.to_numeric(g["fec_corrected"], errors="coerce").dropna()
        baseline_unc = float(first_unc.iloc[0]) if not first_unc.empty else float("nan")
        baseline_corr = float(first_corr.iloc[0]) if not first_corr.empty else float("nan")
        resets = int(iv["counter_reset"].sum()) if not iv.empty else 0

        rows.append({
            "socket_id": sid,
            "test": test,
            "switch": switch,
            "side": g["side"].iloc[0] if "side" in g.columns else None,
            "port": port,
            "port_num": port_short(port),
            "source_file": g["source_file"].iloc[0] if "source_file" in g.columns else None,
            "tz": g["tz"].iloc[0] if "tz" in g.columns else None,
            "fec_mode": (g["fec_mode"].dropna().iloc[0]
                         if "fec_mode" in g.columns and g["fec_mode"].notna().any() else None),
            "samples": int(len(g)),
            "intervals_total": int(len(iv)),
            "intervals_stable": int(len(stable)),
            "intervals_nonstable": int(len(iv) - len(stable)),
            "ts_start": g["ts"].min(),
            "ts_end": g["ts"].max(),
            "span_seconds": ((g["ts"].max() - g["ts"].min()).total_seconds()
                             if len(g) > 1 else 0.0),
            "soak_seconds": stable_seconds,

            # жёсткие счётчики — по всему захвату
            "fec_uncorrected": fec_uncorr_total,
            "fec_corrected": fec_corr_total,
            "fec_uncorrected_stable": unc_attr["STABLE"],
            "fec_uncorrected_flap": unc_attr["FLAP"] + unc_attr["RESET"],
            "fec_uncorrected_linkdown": unc_attr["LINK_DOWN"],
            "carrier_transitions_delta": ct_total,
            "carrier_transitions_stable": ct_attr["STABLE"],
            "link_flaps": link_flaps,
            "link_down_samples": down_samples,
            "defect_samples": defect_samples,

            # скорости — по стабильным интервалам
            "fec_corrected_stable": fec_corr_stable,
            "fec_corrected_rate_eps": corr_rate,
            "fec_corrected_rate_max_eps": rate_max,
            "fec_rate_cv": rate_cv,
            "intervals_with_corrected": int((inst > 0).sum()) if not inst.empty else 0,
            "ber_pre_fec_est": ber_pre_fec,

            **{f"err_{k}": v for k, v in hard_errors.items()},
            "hard_errors_total": float(sum(hard_errors.values())),

            # целостность данных
            "counter_baseline_uncorrected": baseline_unc,
            "counter_baseline_corrected": baseline_corr,
            "counter_zero_based": bool(baseline_unc == 0 and baseline_corr == 0),
            "counter_resets": resets,
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df.sort_values(["test", "switch", "port"], inplace=True, kind="mergesort")
        df.reset_index(drop=True, inplace=True)
    return df



def summarize_dom(df_dom: pd.DataFrame, profile: OpticProfile,
                  th: Optional[Thresholds] = None) -> pd.DataFrame:
    """
    Сводит DOM-телеметрию по каждому сокету.

    Считает по каждому параметру min/max/mean/первое/последнее, дрейф (МНК-наклон,
    нормированный на сутки), перекос между линиями и чувствительность тока
    смещения к температуре — ключевой признак «уставшего» лазера.

    ВАЖНО: статистика считается по снимкам ВНЕ краевых окон. Снимок, сделанный в
    момент, когда оператор извлёк модуль, показывает отсутствие света на приёме
    и, будучи учтённым, породил бы ложную отбраковку по выходу RX за границы
    даташита. На графики при этом выводится полный ряд — чтобы событие было видно.
    """
    if df_dom.empty:
        return pd.DataFrame()

    th = th or Thresholds()
    if "edge" not in df_dom.columns:
        df_dom = df_dom.assign(edge=False)

    rows: List[Dict[str, Any]] = []
    for (test, switch, port), grp in df_dom.groupby(SOCKET_COLS, sort=False):
        full = grp.sort_values("ts", kind="mergesort")
        core = full[~full["edge"].fillna(False).astype(bool)]
        # Если краевые окна съели весь ряд — работаем по полному, но помечаем это.
        g = core if not core.empty else full
        t0 = g["ts"].min()
        hours = (g["ts"] - t0).dt.total_seconds() / 3600.0

        temp = pd.to_numeric(g["temp_c"], errors="coerce")
        # Температура дублируется по 4 линиям — берём по уникальным снимкам.
        temp_by_snap = g.groupby("snapshot")["temp_c"].first()
        temp_series = pd.to_numeric(temp_by_snap, errors="coerce").dropna()

        rec: Dict[str, Any] = {
            "socket_id": socket_id(test, switch, port),
            "test": test,
            "switch": switch,
            "port": port,
            "port_num": port_short(port),
            "dom_samples": int(g["snapshot"].nunique()),
            "dom_samples_excluded": int(full["snapshot"].nunique() - g["snapshot"].nunique()),
            "lanes_seen": int(pd.to_numeric(g["lane"], errors="coerce").nunique()),
            "temp_min_c": float(temp_series.min()) if not temp_series.empty else np.nan,
            "temp_max_c": float(temp_series.max()) if not temp_series.empty else np.nan,
            "temp_mean_c": float(temp_series.mean()) if not temp_series.empty else np.nan,
            "temp_rise_c": (float(temp_series.iloc[-1] - temp_series.iloc[0])
                            if len(temp_series) > 1 else np.nan),
        }

        # --- параметры по линиям -------------------------------------------------
        lane_tx_means, lane_rx_means = [], []
        bias_slopes, tx_slopes, rx_slopes = [], [], []
        tx_all, rx_all, bias_all = [], [], []

        for lane, lg in g.groupby("lane", sort=True):
            lg = lg.sort_values("ts", kind="mergesort")
            lh = (lg["ts"] - t0).dt.total_seconds() / 3600.0
            tx = pd.to_numeric(lg["tx_dbm"], errors="coerce").replace([np.inf, -np.inf], np.nan)
            rx = pd.to_numeric(lg["rx_dbm"], errors="coerce").replace([np.inf, -np.inf], np.nan)
            bias = pd.to_numeric(lg["bias_ma"], errors="coerce")

            lane_key = f"lane{int(lane)}"
            for name, series in (("tx_dbm", tx), ("rx_dbm", rx), ("bias_ma", bias)):
                if series.notna().any():
                    rec[f"{lane_key}_{name}_min"] = float(series.min())
                    rec[f"{lane_key}_{name}_max"] = float(series.max())
                    rec[f"{lane_key}_{name}_mean"] = float(series.mean())
                    rec[f"{lane_key}_{name}_first"] = float(series.dropna().iloc[0])
                    rec[f"{lane_key}_{name}_last"] = float(series.dropna().iloc[-1])
                    slope = linear_slope(lh, series)
                    # Нормируем наклон на сутки.
                    rec[f"{lane_key}_{name}_slope_per_day"] = (
                        slope * 24.0 if slope is not None else np.nan)
                else:
                    for suffix in ("min", "max", "mean", "first", "last", "slope_per_day"):
                        rec[f"{lane_key}_{name}_{suffix}"] = np.nan

            if tx.notna().any():
                lane_tx_means.append(float(tx.mean()))
                tx_all.append(tx)
                s = linear_slope(lh, tx)
                if s is not None:
                    tx_slopes.append(s * 24.0)
            if rx.notna().any():
                lane_rx_means.append(float(rx.mean()))
                rx_all.append(rx)
                s = linear_slope(lh, rx)
                if s is not None:
                    rx_slopes.append(s * 24.0)
            if bias.notna().any():
                bias_all.append(bias)
                s = linear_slope(lh, bias)
                if s is not None:
                    bias_slopes.append(s * 24.0)

        def cat(series_list: List[pd.Series]) -> pd.Series:
            return (pd.concat(series_list) if series_list else pd.Series(dtype=float))

        tx_cat, rx_cat, bias_cat = cat(tx_all), cat(rx_all), cat(bias_all)

        rec.update({
            "tx_dbm_min": float(tx_cat.min()) if not tx_cat.dropna().empty else np.nan,
            "tx_dbm_max": float(tx_cat.max()) if not tx_cat.dropna().empty else np.nan,
            "tx_dbm_mean": float(tx_cat.mean()) if not tx_cat.dropna().empty else np.nan,
            "rx_dbm_min": float(rx_cat.min()) if not rx_cat.dropna().empty else np.nan,
            "rx_dbm_max": float(rx_cat.max()) if not rx_cat.dropna().empty else np.nan,
            "rx_dbm_mean": float(rx_cat.mean()) if not rx_cat.dropna().empty else np.nan,
            "bias_ma_min": float(bias_cat.min()) if not bias_cat.dropna().empty else np.nan,
            "bias_ma_max": float(bias_cat.max()) if not bias_cat.dropna().empty else np.nan,
            "bias_ma_mean": float(bias_cat.mean()) if not bias_cat.dropna().empty else np.nan,

            # Максимальный по модулю дрейф среди линий — «худший случай».
            "tx_drift_db_per_day": (max(tx_slopes, key=abs) if tx_slopes else np.nan),
            "rx_drift_db_per_day": (max(rx_slopes, key=abs) if rx_slopes else np.nan),
            "bias_drift_ma_per_day": (max(bias_slopes, key=abs) if bias_slopes else np.nan),

            # Перекос между линиями (max-min средних значений).
            "tx_lane_imbalance_db": (max(lane_tx_means) - min(lane_tx_means)
                                     if len(lane_tx_means) > 1 else np.nan),
            "rx_lane_imbalance_db": (max(lane_rx_means) - min(lane_rx_means)
                                     if len(lane_rx_means) > 1 else np.nan),
        })

        # --- запас до границ даташита -------------------------------------------
        rec["tx_margin_low_db"] = (rec["tx_dbm_min"] - profile.tx_dbm_min
                                   if math.isfinite(rec["tx_dbm_min"]) else np.nan)
        rec["tx_margin_high_db"] = (profile.tx_dbm_max - rec["tx_dbm_max"]
                                    if math.isfinite(rec["tx_dbm_max"]) else np.nan)
        rec["rx_margin_low_db"] = (rec["rx_dbm_min"] - profile.rx_dbm_min
                                   if math.isfinite(rec["rx_dbm_min"]) else np.nan)
        rec["rx_margin_high_db"] = (profile.rx_dbm_max - rec["rx_dbm_max"]
                                    if math.isfinite(rec["rx_dbm_max"]) else np.nan)

        # --- чувствительность тока смещения к прогреву ---------------------------
        # Здоровый VCSEL/EML со схемой APC слегка увеличивает ток при нагреве.
        # Аномалией считается рост тока, НЕ объяснимый ростом температуры.
        temp_rise = rec.get("temp_rise_c")
        bias_drift = rec.get("bias_drift_ma_per_day")
        if (bias_drift is not None and math.isfinite(bias_drift)
                and temp_rise is not None and math.isfinite(temp_rise)
                and abs(temp_rise) >= 1.0):
            span_days = max((g["ts"].max() - g["ts"].min()).total_seconds() / 86400.0, 1e-6)
            rec["bias_per_degc"] = (bias_drift * span_days) / temp_rise
        else:
            rec["bias_per_degc"] = np.nan

        # --- гейт статистической состоятельности ---------------------------------
        # Наклон, снятый за короткий отрезок и приведённый к суткам, — это
        # умноженный шум. Такие метрики обнуляем в NaN, чтобы движок вердиктов
        # их пропустил, а не выносил приговор по экстраполяции.
        span_hours = ((g["ts"].max() - g["ts"].min()).total_seconds() / 3600.0
                      if len(g) > 1 else 0.0)
        rec["dom_span_hours"] = span_hours
        rec["drift_assessable"] = span_hours >= th.min_hours_for_drift
        if not rec["drift_assessable"]:
            for key in ("tx_drift_db_per_day", "rx_drift_db_per_day",
                        "bias_drift_ma_per_day", "bias_per_degc"):
                rec[key] = np.nan
            for key in list(rec):
                if key.endswith("_slope_per_day"):
                    rec[key] = np.nan

        rec["imbalance_assessable"] = rec["dom_samples"] >= th.min_samples_for_dom_stats
        if not rec["imbalance_assessable"]:
            rec["tx_lane_imbalance_db"] = np.nan
            rec["rx_lane_imbalance_db"] = np.nan

        rows.append(rec)

    df = pd.DataFrame(rows)
    if not df.empty:
        df.sort_values(["test", "switch", "port"], inplace=True, kind="mergesort")
        df.reset_index(drop=True, inplace=True)
    return df


# ======================================================================================
#  8. ДВИЖОК ВЕРДИКТОВ
# ======================================================================================

def _worst(*verdicts: str) -> str:
    """Возвращает самый «тяжёлый» из переданных вердиктов."""
    known = [v for v in verdicts if v in VERDICT_SEVERITY]
    if not known:
        return VERDICT_UNKNOWN
    return max(known, key=lambda v: VERDICT_SEVERITY[v])


def evaluate_socket(
    err: pd.Series,
    dom: Optional[pd.Series],
    profile: OpticProfile,
    th: Thresholds,
) -> Dict[str, Any]:
    """
    Выносит вердикт по одному «посадочному месту» (модуль в конкретном порту
    конкретного круга) и формирует человекочитаемый список причин.

    Возвращает словарь с полями verdict / reasons_fail / reasons_warn / флагами.
    """
    fails: List[str] = []
    warns: List[str] = []
    notes: List[str] = []

    # ---------------------------------------------------------------- FEC ---------
    # Неисправляемые FEC-ошибки — первичный признак браковки. Считаются по ВСЕМУ
    # захвату и не могут быть обнулены ни одним окном или фильтром. Обстоятельства
    # (стабильный прогон / флап / потеря линка) приводятся рядом, чтобы инженер
    # видел контекст, но сам факт наличия ошибок скрыт быть не может.
    uncorr = float(err.get("fec_uncorrected", 0) or 0)
    unc_stable = float(err.get("fec_uncorrected_stable", 0) or 0)
    unc_flap = float(err.get("fec_uncorrected_flap", 0) or 0)
    unc_down = float(err.get("fec_uncorrected_linkdown", 0) or 0)

    if uncorr > th.uncorrected_fail:
        parts = []
        if unc_stable:
            parts.append(f"{fmt_count(unc_stable)} в стабильном прогоне")
        if unc_flap:
            parts.append(f"{fmt_count(unc_flap)} при передёргивании линка/сбросе счётчика")
        if unc_down:
            parts.append(f"{fmt_count(unc_down)} при потере линка")
        detail = f" ({'; '.join(parts)})" if parts else ""
        fails.append(f"FEC Uncorrected = {fmt_count(uncorr)}{detail} — неисправляемые "
                     f"ошибки на канальном уровне")
        if unc_stable > 0:
            fails.append(f"Из них {fmt_count(unc_stable)} набраны при поднятом линке "
                         f"без флапов и сбросов — это свойство тракта, а не "
                         f"следствие монтажных работ")

    ber = err.get("ber_pre_fec_est", float("nan"))
    corr_rate = err.get("fec_corrected_rate_eps", float("nan"))
    if isinstance(ber, float) and math.isfinite(ber):
        if ber > th.ber_warn_max:
            fails.append(f"Оценочный pre-FEC BER = {fmt_ber(ber)} превышает "
                         f"{fmt_ber(th.ber_warn_max)} — нет запаса по бюджету FEC91")
        elif ber > th.ber_pass_max:
            warns.append(f"Оценочный pre-FEC BER = {fmt_ber(ber)} "
                         f"(> {fmt_ber(th.ber_pass_max)}); поток FEC corrected "
                         f"{fmt_num(corr_rate, 1)} ошибок/с — рабочий, но без запаса")

    cv = err.get("fec_rate_cv", 0.0) or 0.0
    if (isinstance(ber, float) and math.isfinite(ber) and ber > th.ber_pass_max
            and cv > th.fec_rate_cv_warn):
        warns.append(f"Поток FEC corrected нестабилен (коэф. вариации {fmt_num(cv)}) "
                     f"— ошибки идут всплесками")

    # ---------------------------------------------------------------- линк --------
    # Флапы также считаются по всему захвату: пропущенный флап — пропущенный брак.
    flaps = int(err.get("link_flaps", 0) or 0)
    if flaps > th.link_flaps_fail:
        fails.append(f"Флапы линка: {flaps} "
                     f"({fmt_count(err.get('carrier_transitions_delta'))} carrier "
                     f"transitions за захват)")
    down = int(err.get("link_down_samples", 0) or 0)
    if down > 0:
        fails.append(f"Зафиксировано {down} снимк(ов) с Physical link is Down")
    defect_n = int(err.get("defect_samples", 0) or 0)
    if defect_n > 0:
        fails.append(f"Зафиксировано {defect_n} снимк(ов) с Active defects != None")

    # -------------------------------------------------- целостность счётчиков -----
    if err.get("counter_zero_based") is False:
        notes.append(
            f"Счётчики не были обнулены перед прогоном (стартовые значения: "
            f"corrected {fmt_count(err.get('counter_baseline_corrected'))}, "
            f"uncorrected {fmt_count(err.get('counter_baseline_uncorrected'))}) — "
            f"в метрику идёт прирост за захват, не абсолютное значение")
    resets = int(err.get("counter_resets", 0) or 0)
    if resets:
        notes.append(f"Обнаружено сбросов счётчика: {resets} "
                     f"(clear interfaces statistics либо пересоздание интерфейса)")

    # ------------------------------------------------- input / CRC / bit errors ---
    for col in HARD_ERROR_COLS:
        val = float(err.get(f"err_{col}", 0) or 0)
        if val > th.input_errors_fail:
            fails.append(f"{col.replace('_', ' ')} = {fmt_count(val)} (> {th.input_errors_fail})")

    # ---------------------------------------------------------------- DOM ---------
    dom_ok = dom is not None and not (isinstance(dom, float) and math.isnan(dom))
    if not dom_ok:
        notes.append("DOM-телеметрия отсутствует — оптические критерии не проверены")
    else:
        def g(key: str) -> float:
            v = dom.get(key, np.nan)
            try:
                return float(v)
            except (TypeError, ValueError):
                return float("nan")

        # Выход за границы даташита -> безусловный брак.
        checks = [
            ("TX power", g("tx_dbm_min"), g("tx_dbm_max"),
             profile.tx_dbm_min, profile.tx_dbm_max, "дБм"),
            ("RX power", g("rx_dbm_min"), g("rx_dbm_max"),
             profile.rx_dbm_min, profile.rx_dbm_max, "дБм"),
            ("Bias current", g("bias_ma_min"), g("bias_ma_max"),
             profile.bias_ma_min, profile.bias_ma_max, "мА"),
            ("Temperature", g("temp_min_c"), g("temp_max_c"),
             profile.temp_c_min, profile.temp_c_max, "°C"),
        ]
        for name, vmin, vmax, lo, hi, unit in checks:
            if math.isfinite(vmin) and vmin < lo:
                fails.append(f"{name} = {fmt_num(vmin)} {unit} ниже границы даташита "
                             f"{profile.name} ({lo} {unit})")
            if math.isfinite(vmax) and vmax > hi:
                fails.append(f"{name} = {fmt_num(vmax)} {unit} выше границы даташита "
                             f"{profile.name} ({hi} {unit})")

        # Малый запас до НИЖНЕЙ границы -> предупреждение: именно он означает
        # отсутствие запаса по бюджету трассы (сигнал у порога чувствительности).
        for label, key in (("TX", "tx_margin_low_db"), ("RX", "rx_margin_low_db")):
            m = g(key)
            if math.isfinite(m) and 0 <= m < th.dom_margin_warn_db:
                warns.append(f"Запас {label} до нижней границы даташита "
                             f"всего {fmt_num(m)} дБ — нет запаса по бюджету трассы")

        # Близость к ВЕРХНЕЙ границе бюджету трассы не угрожает (риск —
        # перегрузка приёмника, и он уже покрыт проверкой rx_dbm_max выше),
        # поэтому это справочное замечание, а не понижение вердикта.
        for label, key in (("TX", "tx_margin_high_db"), ("RX", "rx_margin_high_db")):
            m = g(key)
            if math.isfinite(m) and 0 <= m < th.dom_margin_warn_db:
                notes.append(f"{label} работает вблизи верхней границы даташита "
                             f"(запас {fmt_num(m)} дБ)")

        # Дрейф мощности.
        for label, key in (("TX", "tx_drift_db_per_day"), ("RX", "rx_drift_db_per_day")):
            d = g(key)
            if math.isfinite(d):
                if abs(d) >= th.power_drift_fail_db:
                    fails.append(f"Дрейф {label} мощности {fmt_num(d)} дБ/сутки "
                                 f"(порог {th.power_drift_fail_db}) — деградация")
                elif abs(d) >= th.power_drift_warn_db:
                    warns.append(f"Дрейф {label} мощности {fmt_num(d)} дБ/сутки "
                                 f"(порог {th.power_drift_warn_db})")

        # Перекос между линиями QSFP28.
        for label, key in (("TX", "tx_lane_imbalance_db"), ("RX", "rx_lane_imbalance_db")):
            imb = g(key)
            if math.isfinite(imb):
                if imb >= th.lane_imbalance_fail_db:
                    fails.append(f"Перекос {label} между линиями {fmt_num(imb)} дБ "
                                 f"(порог {th.lane_imbalance_fail_db})")
                elif imb >= th.lane_imbalance_warn_db:
                    warns.append(f"Перекос {label} между линиями {fmt_num(imb)} дБ "
                                 f"(порог {th.lane_imbalance_warn_db})")

        # Если наблюдение слишком короткое, трендовые критерии не применялись —
        # об этом нужно сказать прямо, иначе «в ЗИП» будет выглядеть проверенным
        # по всем пунктам, тогда как часть проверок просто не выполнялась.
        if dom.get("drift_assessable") is False:
            notes.append(f"Дрейф TX/RX/Bias не оценивался: длительность наблюдения "
                         f"{fmt_num(dom.get('dom_span_hours'), 1)} ч "
                         f"(< {th.min_hours_for_drift:g} ч)")
        if dom.get("imbalance_assessable") is False:
            notes.append(f"Перекос между линиями не оценивался: снимков DOM "
                         f"{fmt_count(dom.get('dom_samples'))} "
                         f"(< {th.min_samples_for_dom_stats})")

        # Рост тока смещения — «уставший» лазер.
        bd = g("bias_drift_ma_per_day")
        if math.isfinite(bd):
            if bd >= th.bias_drift_fail_ma_per_day:
                fails.append(f"Лавинообразный рост Bias current {fmt_num(bd)} мА/сутки "
                             f"(порог {th.bias_drift_fail_ma_per_day}) — уставший лазер")
            elif bd >= th.bias_drift_warn_ma_per_day:
                warns.append(f"Рост Bias current {fmt_num(bd)} мА/сутки "
                             f"(порог {th.bias_drift_warn_ma_per_day})")
        bpd = g("bias_per_degc")
        if math.isfinite(bpd) and abs(bpd) >= th.bias_per_degc_warn:
            warns.append(f"Ток смещения растёт на {fmt_num(bpd)} мА/°C — "
                         f"повышенная чувствительность к прогреву")
        bmax = g("bias_ma_max")
        if math.isfinite(bmax) and bmax > profile.bias_ma_typ_max:
            warns.append(f"Bias current {fmt_num(bmax)} мА выше типового значения "
                         f"{profile.bias_ma_typ_max} мА для {profile.name}")

    # -------------------------------------------------- достаточность выборки -----
    usable = int(err.get("intervals_stable", 0) or 0)
    soak_h = float(err.get("soak_seconds", 0) or 0) / 3600.0
    insufficient = usable == 0 or soak_h < th.min_hours_for_pass
    if usable == 0:
        notes.append("Нет ни одного стабильного интервала — метрики надёжности "
                     "не рассчитаны")
    elif insufficient:
        notes.append(f"Наработка в стабильном режиме всего {soak_h:.1f} ч "
                     f"(< {th.min_hours_for_pass:g} ч) — недостаточно для аттестации; "
                     f"отсутствие ошибок за такой срок ничего не доказывает")

    # Порядок важен: найденные дефекты остаются дефектами и при короткой
    # выборке, а вот ПОЛОЖИТЕЛЬНЫЙ вердикт короткой выборкой не обосновывается.
    if fails:
        verdict = VERDICT_FAIL
    elif warns:
        verdict = VERDICT_WARN
    elif insufficient:
        verdict = VERDICT_UNKNOWN
    else:
        verdict = VERDICT_PASS

    return {
        "verdict": verdict,
        "reasons_fail": fails,
        "reasons_warn": warns,
        "notes": notes,
        "reason_text": " | ".join(fails + warns + notes) or "Все критерии в норме",
    }


def evaluate_all_sockets(
    df_err: pd.DataFrame,
    df_dom: pd.DataFrame,
    profile: OpticProfile,
    th: Thresholds,
) -> pd.DataFrame:
    """Прогоняет движок вердиктов по всем сокетам и склеивает результат с метриками."""
    if df_err.empty:
        return pd.DataFrame()

    dom_idx = (df_dom.set_index("socket_id") if not df_dom.empty
               and "socket_id" in df_dom.columns else pd.DataFrame())

    results: List[Dict[str, Any]] = []
    for _, err in df_err.iterrows():
        sid = err["socket_id"]
        dom_row = dom_idx.loc[sid] if (not dom_idx.empty and sid in dom_idx.index) else None
        if isinstance(dom_row, pd.DataFrame):  # на случай дублей индекса
            dom_row = dom_row.iloc[0]
        try:
            res = evaluate_socket(err, dom_row, profile, th)
        except Exception as exc:  # noqa: BLE001
            LOG.error("Сбой оценки сокета %s: %s", sid, exc)
            res = {"verdict": VERDICT_UNKNOWN, "reasons_fail": [], "reasons_warn": [],
                   "notes": [f"Ошибка расчёта: {exc}"],
                   "reason_text": f"Ошибка расчёта: {exc}"}
        results.append({"socket_id": sid, **res})

    df_res = pd.DataFrame(results)
    out = df_err.merge(df_res, on="socket_id", how="left")
    return out


# ======================================================================================
#  9. ЛОГИКА РОТАЦИИ В 2 КРУГА
# ======================================================================================

ROT_CONFIRMED_FAIL = "БРАК подтверждён в двух кругах"
ROT_LINK_FAULT = "Годен: виновата трасса/патч-корд (чист в другом круге)"
ROT_NEEDS_RETEST = "Требуется повторный тест (данных одного круга недостаточно)"
ROT_CLEAN_BOTH = "Чист во всех кругах"
ROT_SINGLE_CLEAN = "Чист (один круг)"
ROT_NO_MAP = "Кросс-раундовый анализ недоступен (нет карты S/N)"


def build_link_view(df_sockets: pd.DataFrame, module_map: ModuleMap) -> pd.DataFrame:
    """
    Сводит сокеты в линки: для каждого сокета находит партнёра на другом конце
    и склеивает вердикты обеих сторон.
    """
    if df_sockets.empty:
        return pd.DataFrame()

    rows: List[Dict[str, Any]] = []
    seen: set = set()
    for _, r in df_sockets.iterrows():
        partner = resolve_partner(df_sockets, r["test"], r["switch"], r["port"], module_map)
        if partner is None:
            key = (r["socket_id"], None)
            if key in seen:
                continue
            seen.add(key)
            rows.append({
                "test": r["test"],
                "link": f"{r['switch']}:{r['port']} ↔ (партнёр не найден)",
                "a_socket": r["socket_id"], "a_serial": r.get("serial"),
                "a_verdict": r["verdict"],
                "b_socket": None, "b_serial": None, "b_verdict": None,
                "link_verdict": r["verdict"],
                "link_dirty": r["verdict"] == VERDICT_FAIL,
            })
            continue

        p_sid = socket_id(*partner)
        pair = tuple(sorted([r["socket_id"], p_sid]))
        if pair in seen:
            continue
        seen.add(pair)

        prow = df_sockets[df_sockets["socket_id"] == p_sid]
        if prow.empty:
            continue
        prow = prow.iloc[0]

        # Сторону A ставим первой для стабильного порядка.
        a, b = (r, prow) if str(r["switch"]) <= str(prow["switch"]) else (prow, r)
        link_verdict = _worst(a["verdict"], b["verdict"])
        rows.append({
            "test": a["test"],
            "link": f"{a['switch']}:{port_short(a['port'])} ↔ "
                    f"{b['switch']}:{port_short(b['port'])}",
            "a_socket": a["socket_id"], "a_serial": a.get("serial"),
            "a_verdict": a["verdict"],
            "b_socket": b["socket_id"], "b_serial": b.get("serial"),
            "b_verdict": b["verdict"],
            "a_fec_uncorrected": a.get("fec_uncorrected"),
            "b_fec_uncorrected": b.get("fec_uncorrected"),
            "a_ber": a.get("ber_pre_fec_est"),
            "b_ber": b.get("ber_pre_fec_est"),
            "a_flaps": a.get("link_flaps"),
            "b_flaps": b.get("link_flaps"),
            "link_verdict": link_verdict,
            "link_dirty": link_verdict == VERDICT_FAIL,
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df.sort_values(["test", "link"], inplace=True, kind="mergesort")
        df.reset_index(drop=True, inplace=True)
    return df


def apply_rotation_logic(
    df_sockets: pd.DataFrame, module_map: ModuleMap,
) -> pd.DataFrame:
    """
    Агрегирует вердикты по серийному номеру через все круги ротации.

    Правила:
      * FAIL в >= 2 кругах  -> подтверждённый брак (виноват сам модуль);
      * FAIL в 1 круге при наличии чистого результата в другом -> виновата трасса
        (патч-корд/загрязнение коннектора), модуль годен после перечистки;
      * FAIL в единственном круге, второго прогона нет -> требуется повторный тест;
      * чист везде -> вердикт по худшему из кругов (PASS или WARNING).
    """
    if df_sockets.empty:
        return pd.DataFrame()

    rows: List[Dict[str, Any]] = []
    for serial, grp in df_sockets.groupby("serial", sort=False):
        tests = sorted(set(grp["test"]))
        verdicts = list(grp["verdict"])
        fail_tests = sorted(set(grp.loc[grp["verdict"] == VERDICT_FAIL, "test"]))
        clean_tests = sorted(set(grp.loc[grp["verdict"].isin([VERDICT_PASS, VERDICT_WARN]),
                                         "test"]))
        worst = _worst(*verdicts)
        has_map = bool(grp["serial_is_real"].any()) if "serial_is_real" in grp.columns else False

        if len(tests) < 2:
            if not has_map:
                rotation = ROT_NO_MAP
            elif worst == VERDICT_FAIL:
                rotation = ROT_NEEDS_RETEST
            else:
                rotation = ROT_SINGLE_CLEAN
            final = worst
        else:
            if len(fail_tests) >= 2:
                rotation = ROT_CONFIRMED_FAIL
                final = VERDICT_FAIL
            elif len(fail_tests) == 1 and clean_tests:
                rotation = ROT_LINK_FAULT
                # Модуль реабилитирован, но с пометкой «не в ЗИП» — он уже
                # отработал в проблемном линке и требует перечистки/перепроверки.
                final = VERDICT_WARN
            elif fail_tests:
                rotation = ROT_NEEDS_RETEST
                final = VERDICT_FAIL
            else:
                rotation = ROT_CLEAN_BOTH
                final = worst

        reasons = []
        for _, r in grp.iterrows():
            if r["verdict"] in (VERDICT_FAIL, VERDICT_WARN):
                reasons.append(f"[Круг {r['test']} · {r['switch']}:{port_short(r['port'])}] "
                               f"{r['reason_text']}")

        rows.append({
            "serial": serial,
            "serial_is_real": has_map,
            "tests_participated": ", ".join(f"Круг {t}" for t in tests),
            "n_tests": len(tests),
            "sockets": " ; ".join(f"T{r['test']}:{r['switch']}:{port_short(r['port'])}"
                                  for _, r in grp.iterrows()),
            "verdict_per_test": " ; ".join(f"К{r['test']}={r['verdict']}"
                                           for _, r in grp.iterrows()),
            "rotation_conclusion": rotation,
            "final_verdict": final,
            "fec_uncorrected_total": float(pd.to_numeric(
                grp.get("fec_uncorrected", pd.Series(dtype=float)),
                errors="coerce").fillna(0).sum()),
            "link_flaps_total": int(pd.to_numeric(
                grp.get("link_flaps", pd.Series(dtype=float)),
                errors="coerce").fillna(0).sum()),
            "ber_worst": (float(pd.to_numeric(grp.get("ber_pre_fec_est",
                                                      pd.Series(dtype=float)),
                                              errors="coerce").max())
                          if "ber_pre_fec_est" in grp.columns else np.nan),
            "soak_hours_total": float(pd.to_numeric(
                grp.get("soak_seconds", pd.Series(dtype=float)),
                errors="coerce").fillna(0).sum()) / 3600.0,
            "reason_text": " || ".join(reasons) or "Все критерии в норме во всех кругах",
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["_sev"] = df["final_verdict"].map(VERDICT_SEVERITY).fillna(0)
        df.sort_values(["_sev", "serial"], ascending=[False, True], inplace=True,
                       kind="mergesort")
        df.drop(columns=["_sev"], inplace=True)
        df.reset_index(drop=True, inplace=True)
    return df


# ======================================================================================
#  10. ВИЗУАЛИЗАЦИЯ
# ======================================================================================
#
#  Палитра и правила оформления соответствуют проверенной дизайн-системе:
#    * категориальные цвета назначаются по фиксированным слотам, без циклирования;
#    * ни одного графика с двумя осями Y (разные величины -> отдельные панели);
#    * тонкие штрихи, сплошная волосяная сетка, приглушённые оси;
#    * легенда присутствует всегда при >= 2 рядах, точечные подписи — выборочно;
#    * цвета статуса (годен/предупреждение/брак) зарезервированы и всегда
#      сопровождаются текстовой меткой, то есть смысл не несётся одним лишь цветом.
#
#  Палитра из 4 слотов проверена валидатором (светлая подложка #fcfcfb):
#    lightness band PASS, chroma floor PASS, CVD ΔE 9.1 PASS, normal-vision ΔE 22.9 PASS.
# --------------------------------------------------------------------------------------

#: Категориальные слоты (порядок фиксирован) — используются для 4 линий QSFP28.
SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]
#: Маркеры — вторичное кодирование идентичности линии (для ч/б печати и CVD).
SERIES_MARKERS = ["o", "s", "^", "D"]

#: Зарезервированные цвета статуса.
STATUS_GOOD = "#0ca30c"
STATUS_WARNING = "#fab219"
STATUS_SERIOUS = "#ec835a"
STATUS_CRITICAL = "#d03b3b"

STATUS_BY_VERDICT = {
    VERDICT_PASS: STATUS_GOOD,
    VERDICT_WARN: STATUS_WARNING,
    VERDICT_FAIL: STATUS_CRITICAL,
    VERDICT_UNKNOWN: "#898781",
}

#: Хром графика.
SURFACE = "#fcfcfb"
INK_PRIMARY = "#0b0b0b"
INK_SECONDARY = "#52514e"
INK_MUTED = "#898781"
GRIDLINE = "#e1e0d9"
BASELINE = "#c3c2b7"


def setup_matplotlib() -> None:
    """Единая настройка стиля для всех графиков отчёта."""
    plt.rcParams.update({
        "figure.facecolor": SURFACE,
        "axes.facecolor": SURFACE,
        "savefig.facecolor": SURFACE,
        "font.family": "DejaVu Sans",          # содержит кириллицу
        "font.size": 9,
        "axes.titlesize": 11,
        "axes.titleweight": "bold",
        "axes.titlecolor": INK_PRIMARY,
        "axes.labelsize": 9,
        "axes.labelcolor": INK_SECONDARY,
        "axes.edgecolor": BASELINE,
        "axes.linewidth": 0.8,
        "axes.grid": True,
        "axes.axisbelow": True,
        "grid.color": GRIDLINE,
        "grid.linewidth": 0.6,
        "grid.linestyle": "-",                 # сплошная волосяная сетка, не пунктир
        "xtick.color": INK_MUTED,
        "ytick.color": INK_MUTED,
        "xtick.labelsize": 8,
        "ytick.labelsize": 8,
        "legend.frameon": False,
        "legend.fontsize": 8,
        "lines.linewidth": 1.6,
        "lines.markersize": 3.2,
        "figure.dpi": 150,
    })


def _despine(ax: plt.Axes) -> None:
    """Убирает верхнюю и правую рамки — сетка и так задаёт систему координат."""
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)


def _time_axis(ax: plt.Axes, hours_span: float) -> None:
    """Настраивает ось времени под длительность прогона."""
    if hours_span > 18:
        loc = mdates.HourLocator(interval=4)
    elif hours_span > 6:
        loc = mdates.HourLocator(interval=2)
    else:
        loc = mdates.HourLocator(interval=1)
    ax.xaxis.set_major_locator(loc)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m\n%H:%M"))


def _thousands(x: float, _pos: int = 0) -> str:
    """Разделитель тысяч неразрывным пробелом для осей."""
    if abs(x) >= 1000:
        return f"{x:,.0f}".replace(",", " ")
    if abs(x) >= 1:
        return f"{x:.0f}"
    return f"{x:g}"


def _series_legend(ax: plt.Axes, ncol: int = 4) -> None:
    """
    Ставит легенду НАД полем графика, а не поверх данных.

    Заголовок панели выровнен влево, поэтому правый верхний угол свободен —
    легенда там не перекрывает ни линии, ни опорные подписи.
    """
    if not ax.get_legend_handles_labels()[0]:
        return
    ax.legend(loc="lower right", bbox_to_anchor=(1.0, 1.0), ncol=ncol,
              handlelength=1.6, columnspacing=1.3, borderaxespad=0.0)


def _frame_with_refs(
    ax: plt.Axes,
    values: pd.Series,
    refs: Sequence[Tuple[float, str, str]],
    unit: str,
    min_pad: float = 0.4,
) -> None:
    """
    Масштабирует ось Y по данным и наносит опорные линии (границы даташита)
    ТОЛЬКО если они попадают в разумную окрестность данных.

    Зачем: жёсткое включение далёких границ даташита в поле зрения сплющивает
    реальные данные в тонкую полоску, и весь смысл графика — увидеть дрейф —
    теряется. Если граница далеко, она выносится текстовой сноской в углу:
    информация сохраняется, а разрешение графика — тоже.

    refs: последовательность кортежей (уровень, подпись, цвет).
    """
    vals = (pd.to_numeric(values, errors="coerce")
            .replace([np.inf, -np.inf], np.nan).dropna())
    if vals.empty:
        return

    dmin, dmax = float(vals.min()), float(vals.max())
    span = max(dmax - dmin, min_pad)
    pad = max(span * 0.32, min_pad)
    lo_lim, hi_lim = dmin - pad, dmax + pad
    # Опорная линия «близка», если отстоит от данных не более чем на их размах.
    tolerance = max(span * 1.2, min_pad * 2)

    far: List[str] = []
    for level, label, color in refs:
        if level is None or not math.isfinite(level):
            continue
        if (dmin - tolerance) <= level <= (dmax + tolerance):
            ax.axhline(level, color=color, linewidth=1.0, zorder=1)
            ax.annotate(f"{label} {level:g} {unit}",
                        xy=(0.004, level), xycoords=("axes fraction", "data"),
                        ha="left", va="bottom", fontsize=7, color=INK_MUTED)
            lo_lim = min(lo_lim, level - pad * 0.4)
            hi_lim = max(hi_lim, level + pad * 0.4)
        else:
            far.append(f"{label} {level:g} {unit}")

    ax.set_ylim(lo_lim, hi_lim)
    if far:
        ax.annotate("вне поля графика: " + "; ".join(far),
                    xy=(0.996, 0.04), xycoords="axes fraction",
                    ha="right", va="bottom", fontsize=6.8, color=INK_MUTED)


def _note_identical_lanes(ax: plt.Axes, dom: pd.DataFrame, column: str) -> bool:
    """
    Если модуль отдаёт одинаковые значения по всем линиям, кривые на графике
    сливаются в одну. Помечаем это явно — иначе читатель решит, что данные потеряны.
    """
    try:
        per_lane = dom.groupby("lane")[column].apply(
            lambda s: tuple(pd.to_numeric(s, errors="coerce").round(4).tolist()))
        if len(per_lane) > 1 and len(set(per_lane)) == 1:
            ax.annotate("модуль рапортует одинаковое значение по всем линиям — "
                        "кривые совпадают",
                        xy=(0.996, 0.06), xycoords="axes fraction",
                        ha="right", va="bottom", fontsize=6.8, color=INK_MUTED)
            return True
    except Exception:  # noqa: BLE001
        pass
    return False


def _shade_edges(ax: plt.Axes, frame: pd.DataFrame) -> bool:
    """
    Затеняет краевые окна (монтаж/демонтаж) на графиках DOM.

    Данные в этих окнах выводятся — событие должно быть видно, — но в расчёт
    статистики и вердиктов они не идут, и читатель обязан это различать.
    Возвращает True, если хотя бы одно окно затенено.
    """
    if frame.empty or "edge" not in frame.columns:
        return False
    marked = frame[frame["edge"].fillna(False).astype(bool)]
    if marked.empty:
        return False
    ts = marked["ts"].sort_values().unique()
    # Склеиваем соседние метки в непрерывные полосы.
    spans, start, prev = [], ts[0], ts[0]
    step = pd.Timedelta(minutes=12)
    for t in ts[1:]:
        if pd.Timestamp(t) - pd.Timestamp(prev) > step:
            spans.append((start, prev))
            start = t
        prev = t
    spans.append((start, prev))
    for a, b in spans:
        ax.axvspan(a, b, color=GRIDLINE, alpha=0.9, linewidth=0, zorder=0)
    return True


def chart_thermal(
    dom: pd.DataFrame, title: str, profile: OpticProfile, out_path: Path,
) -> Optional[Path]:
    """
    Прогрев модуля: температура и ток смещения по 4 линиям.

    Две ВЕРТИКАЛЬНО РАЗНЕСЁННЫЕ панели с общей осью времени — намеренно не
    совмещённые оси Y: температура (°C) и ток (мА) несопоставимы по масштабу,
    а совмещение двух шкал на одном поле создаёт ложную корреляцию.
    """
    if dom.empty:
        return None
    try:
        fig, (ax_t, ax_b) = plt.subplots(
            2, 1, figsize=(9.6, 4.1), sharex=True,
            gridspec_kw={"height_ratios": [1.0, 1.35], "hspace": 0.22})

        # Краевые окна показываем, но статистику по ним не считаем.
        core = (dom[~dom["edge"].fillna(False).astype(bool)]
                if "edge" in dom.columns else dom)
        if core.empty:
            core = dom

        # --- панель 1: температура модуля (один ряд -> легенда не нужна) ---------
        temp = (dom.groupby("ts")["temp_c"].first().dropna().sort_index())
        temp_core = (core.groupby("ts")["temp_c"].first().dropna().sort_index())
        if not temp.empty:
            _shade_edges(ax_t, dom)
            ax_t.plot(temp.index, temp.values, color=SERIES_COLORS[0],
                      linewidth=1.6, zorder=3)
            ax_t.margins(y=0.16)
            if not temp_core.empty:
                # Выборочная прямая подпись: финальное значение прогона (без
                # краевого окна, где модуль уже извлечён из клетки).
                ax_t.annotate(f"{temp_core.iloc[-1]:.0f} °C",
                              xy=(temp_core.index[-1], temp_core.iloc[-1]),
                              xytext=(5, 0), textcoords="offset points",
                              va="center", fontsize=8, color=INK_SECONDARY)
                rise = temp_core.iloc[-1] - temp_core.iloc[0]
                ax_t.set_title(
                    f"Прогрев: температура модуля (старт {temp_core.iloc[0]:.0f} °C → "
                    f"финиш {temp_core.iloc[-1]:.0f} °C, Δ {rise:+.0f} °C)", loc="left")
        ax_t.set_ylabel("Температура, °C")
        _despine(ax_t)

        # --- панель 2: ток смещения по линиям (4 ряда -> легенда обязательна) ----
        _shade_edges(ax_b, dom)
        lanes = sorted(pd.to_numeric(dom["lane"], errors="coerce").dropna().unique())
        for i, lane in enumerate(lanes[:8]):
            lg = dom[dom["lane"] == lane].sort_values("ts")
            series = pd.to_numeric(lg["bias_ma"], errors="coerce")
            if not series.notna().any():
                continue
            color = SERIES_COLORS[i % len(SERIES_COLORS)]
            marker = SERIES_MARKERS[i % len(SERIES_MARKERS)]
            step = max(len(lg) // 22, 1)   # маркеры для вторичного кодирования
            ax_b.plot(lg["ts"], series, color=color, linewidth=1.5,
                      marker=marker, markevery=step, markersize=3.0,
                      markeredgecolor=SURFACE, markeredgewidth=0.6,
                      label=f"Линия {int(lane)}")
        ax_b.set_ylabel("Ток смещения, мА")
        ax_b.set_title("Ток смещения лазера по линиям", loc="left")
        _frame_with_refs(
            ax_b, core["bias_ma"],
            [(profile.bias_ma_typ_max, "типовой максимум", STATUS_WARNING),
             (profile.bias_ma_max, "предел по даташиту", STATUS_CRITICAL)],
            unit="мА", min_pad=0.15)
        _note_identical_lanes(ax_b, core, "bias_ma")
        _series_legend(ax_b, ncol=min(len(lanes), 4))
        _despine(ax_b)

        span_h = ((dom["ts"].max() - dom["ts"].min()).total_seconds() / 3600.0
                  if len(dom) > 1 else 1.0)
        _time_axis(ax_b, span_h)
        fig.suptitle(title, fontsize=11, fontweight="bold",
                     color=INK_PRIMARY, x=0.008, ha="left", y=0.995)
        fig.subplots_adjust(top=0.90, left=0.085, right=0.975, bottom=0.13)
        fig.savefig(out_path, bbox_inches="tight")
        plt.close(fig)
        return out_path
    except Exception as exc:  # noqa: BLE001
        LOG.error("Не удалось построить график прогрева (%s): %s", title, exc)
        plt.close("all")
        return None


def chart_optical_power(
    dom: pd.DataFrame, title: str, profile: OpticProfile, out_path: Path,
) -> Optional[Path]:
    """
    Мощность TX и RX по каждой из 4 линий — две панели с общей осью времени.

    Обе величины в дБм, но это РАЗНЫЕ измерения (передача и приём), поэтому они
    разнесены по панелям, а не наложены: наложение мешает увидеть перекос
    внутри каждой группы. Границы даташита нанесены как опорные линии.
    """
    if dom.empty:
        return None
    try:
        fig, (ax_tx, ax_rx) = plt.subplots(
            2, 1, figsize=(9.6, 4.5), sharex=True, gridspec_kw={"hspace": 0.26})

        lanes = sorted(pd.to_numeric(dom["lane"], errors="coerce").dropna().unique())
        # Краевые окна рисуем, но не даём им задавать масштаб и статистику:
        # в момент извлечения модуля RX проваливается на десятки дБ.
        core = (dom[~dom["edge"].fillna(False).astype(bool)]
                if "edge" in dom.columns else dom)
        if core.empty:
            core = dom

        for ax, col, label, lo, hi in (
            (ax_tx, "tx_dbm", "TX передача", profile.tx_dbm_min, profile.tx_dbm_max),
            (ax_rx, "rx_dbm", "RX приём", profile.rx_dbm_min, profile.rx_dbm_max),
        ):
            _shade_edges(ax, dom)
            means: List[Tuple[int, float]] = []
            for i, lane in enumerate(lanes[:8]):
                lg = dom[dom["lane"] == lane].sort_values("ts")
                series = (pd.to_numeric(lg[col], errors="coerce")
                          .replace([np.inf, -np.inf], np.nan))
                if not series.notna().any():
                    continue
                color = SERIES_COLORS[i % len(SERIES_COLORS)]
                marker = SERIES_MARKERS[i % len(SERIES_MARKERS)]
                step = max(len(lg) // 22, 1)
                ax.plot(lg["ts"], series, color=color, linewidth=1.5,
                        marker=marker, markevery=step, markersize=3.0,
                        markeredgecolor=SURFACE, markeredgewidth=0.6,
                        label=f"Линия {int(lane)}")
                # Перекос считаем по «чистой» части — так же, как в сводке.
                core_lane = (pd.to_numeric(core.loc[core["lane"] == lane, col],
                                           errors="coerce")
                             .replace([np.inf, -np.inf], np.nan))
                if core_lane.notna().any():
                    means.append((int(lane), float(core_lane.mean())))

            imbalance = (max(m for _, m in means) - min(m for _, m in means)
                         if len(means) > 1 else 0.0)
            # Заголовок держим коротким: справа от него встаёт легенда.
            ax.set_title(f"{label}, перекос линий {imbalance:.2f} дБ", loc="left")
            ax.set_ylabel("Мощность, дБм")

            # Масштаб — по данным; границы даташита наносятся, только если
            # попадают в окрестность данных, иначе уходят в текстовую сноску.
            _frame_with_refs(
                ax, core[col],
                [(lo, "мин. по даташиту", STATUS_SERIOUS),
                 (hi, "макс. по даташиту", STATUS_SERIOUS)],
                unit="дБм", min_pad=0.5)
            _note_identical_lanes(ax, core, col)
            _series_legend(ax, ncol=min(len(lanes), 4))
            _despine(ax)

        span_h = ((dom["ts"].max() - dom["ts"].min()).total_seconds() / 3600.0
                  if len(dom) > 1 else 1.0)
        _time_axis(ax_rx, span_h)
        fig.suptitle(title, fontsize=11, fontweight="bold",
                     color=INK_PRIMARY, x=0.008, ha="left", y=0.995)
        fig.subplots_adjust(top=0.90, left=0.085, right=0.975, bottom=0.13)
        fig.savefig(out_path, bbox_inches="tight")
        plt.close(fig)
        return out_path
    except Exception as exc:  # noqa: BLE001
        LOG.error("Не удалось построить график мощности (%s): %s", title, exc)
        plt.close("all")
        return None


def chart_fec_timeline(
    iv: pd.DataFrame, title: str, out_path: Path,
) -> Optional[Path]:
    """
    FEC во времени: скорость исправленных ошибок и события неисправляемых.

    ВСЕ данные наносятся на график, включая нештатные интервалы — скрывать
    ошибки нельзя. Нештатные участки (флап линка, сброс счётчика, потеря
    сигнала) затенены серым: это подсказка о ПРИЧИНЕ, а не повод убрать
    значение с графика. Каждое событие неисправляемых ошибок отмечается
    отдельной маркой с подписью величины.
    """
    if iv.empty:
        return None
    try:
        fig, (ax_c, ax_u) = plt.subplots(
            2, 1, figsize=(9.6, 3.9), sharex=True,
            gridspec_kw={"height_ratios": [1.2, 1.0], "hspace": 0.30})

        g = iv.sort_values("ts_start").reset_index(drop=True)
        dur = g["duration_s"].replace(0, np.nan)
        d_corr = pd.to_numeric(g["delta_fec_corrected"], errors="coerce").clip(lower=0)
        d_unc = pd.to_numeric(g["delta_fec_uncorrected"], errors="coerce").clip(lower=0)
        rate_c = (d_corr / dur).replace([np.inf, -np.inf], np.nan)
        stable = (g["stable"].fillna(False).astype(bool)
                  if "stable" in g.columns else pd.Series(True, index=g.index))

        # Затеняем нештатные интервалы — как контекст, а не как исключение.
        for _, row in g[~stable].iterrows():
            for ax in (ax_c, ax_u):
                ax.axvspan(row["ts_start"], row["ts_end"],
                           color=GRIDLINE, alpha=0.9, linewidth=0, zorder=0)

        # --- панель 1: скорость исправленных ошибок, ВЕСЬ ряд -------------------
        ax_c.plot(g["ts_start"], rate_c, color=SERIES_COLORS[0],
                  linewidth=1.5, zorder=3)
        if (rate_c.dropna() > 0).any():
            ax_c.set_yscale("symlog", linthresh=1.0)
        ax_c.margins(y=0.25)
        peak = rate_c.max()
        if pd.notna(peak) and peak > 0:
            pk = rate_c.idxmax()
            ax_c.annotate(f"пик {_thousands(float(peak))} ош./с",
                          xy=(g.loc[pk, "ts_start"], float(peak)),
                          xytext=(7, -9), textcoords="offset points",
                          ha="left", va="top", fontsize=7.5, color=INK_SECONDARY)
        ax_c.yaxis.set_major_formatter(FuncFormatter(_thousands))
        ax_c.set_ylabel("FEC corrected,\nошибок/с")
        ax_c.set_title("Скорость исправленных FEC-ошибок "
                       "(серым — нештатные интервалы: флап, сброс счётчика, потеря линка)",
                       loc="left")
        _despine(ax_c)

        # --- панель 2: события неисправляемых ошибок ---------------------------
        total_u = float(d_unc.sum())
        stable_u = float(d_unc[stable].sum())
        events = g.index[d_unc > 0]

        if total_u > 0:
            # Ступенчатое накопление показывает, когда именно «прилетело».
            ax_u.step(g["ts_end"], d_unc.cumsum(), where="post",
                      color=INK_MUTED, linewidth=1.2, zorder=2,
                      label="накоплено за захват")
            for idx in events:
                is_stable = bool(stable.iloc[idx])
                color = STATUS_CRITICAL if is_stable else STATUS_WARNING
                ax_u.plot([g.loc[idx, "ts_end"]], [d_unc.cumsum().iloc[idx]],
                          marker="o" if is_stable else "^", markersize=6,
                          color=color, markeredgecolor=SURFACE, markeredgewidth=0.8,
                          zorder=4, linestyle="none")
                ax_u.annotate(f"+{_thousands(float(d_unc.iloc[idx]))}",
                              xy=(g.loc[idx, "ts_end"], d_unc.cumsum().iloc[idx]),
                              xytext=(0, 7), textcoords="offset points",
                              ha="center", fontsize=7, color=color, zorder=5)
            ax_u.margins(y=0.30)
            ax_u.yaxis.set_major_formatter(FuncFormatter(_thousands))
            note = (f"из них {_thousands(stable_u)} в стабильном прогоне"
                    if stable_u > 0 else
                    "все — в нештатных интервалах (флап / сброс / потеря линка)")
            ax_u.set_title(f"Неисправляемые FEC-ошибки: всего {_thousands(total_u)}, "
                           f"{note}", loc="left")
            # Легенда форм: круг — стабильный прогон, треугольник — нештатный.
            ax_u.plot([], [], marker="o", linestyle="none", color=STATUS_CRITICAL,
                      markersize=6, label="в стабильном прогоне")
            ax_u.plot([], [], marker="^", linestyle="none", color=STATUS_WARNING,
                      markersize=6, label="при флапе / сбросе / потере линка")
            ax_u.legend(loc="upper left", ncol=3, handlelength=1.2,
                        columnspacing=1.0, fontsize=7)
        else:
            ax_u.set_ylim(-0.05, 1.0)
            ax_u.set_title("Неисправляемые FEC-ошибки: не зафиксировано", loc="left")
            ax_u.annotate("за весь период наблюдения — ни одной", xy=(0.5, 0.5),
                          xycoords="axes fraction", ha="center", va="center",
                          fontsize=9, color=STATUS_GOOD)
        ax_u.set_ylabel("FEC uncorrected,\nнакопленным итогом")
        _despine(ax_u)

        span_h = ((g["ts_end"].max() - g["ts_start"].min()).total_seconds() / 3600.0
                  if len(g) > 1 else 1.0)
        _time_axis(ax_u, span_h)
        fig.suptitle(title, fontsize=11, fontweight="bold",
                     color=INK_PRIMARY, x=0.008, ha="left", y=0.995)
        fig.subplots_adjust(top=0.88, left=0.10, right=0.975, bottom=0.15)
        fig.savefig(out_path, bbox_inches="tight")
        plt.close(fig)
        return out_path
    except Exception as exc:  # noqa: BLE001
        LOG.error("Не удалось построить график FEC (%s): %s", title, exc)
        plt.close("all")
        return None


def chart_ber_overview(
    df_sockets: pd.DataFrame, th: Thresholds, out_path: Path,
) -> Optional[Path]:
    """
    Обзорная панель: оценочный pre-FEC BER по всем сокетам, один ряд — один цвет.

    Логарифмическая шкала: значения различаются на порядки. Пороги «эталонно
    чисто» и «нет запаса» нанесены опорными линиями с подписями.
    """
    if df_sockets.empty:
        return None
    try:
        d = df_sockets.copy()
        d["ber"] = pd.to_numeric(d["ber_pre_fec_est"], errors="coerce").fillna(0.0)
        d["label"] = d.apply(
            lambda r: f"К{r['test']} · {r['switch']}:{port_short(r['port'])}", axis=1)
        d = d.sort_values("ber", ascending=True)

        # Ноль на логарифмической шкале непредставим. Рисовать для него столбик
        # «на полу» — значит показать величину там, где её нет, поэтому нулевые
        # места остаются без марки, а смысл несёт текстовая подпись.
        nonzero = d.loc[d["ber"] > 0, "ber"]
        x_lo = (float(nonzero.min()) / 8.0 if not nonzero.empty
                else th.ber_pass_max / 100.0)
        x_hi = max(float(nonzero.max()) * 40.0 if not nonzero.empty else 0.0,
                   th.ber_warn_max * 8.0)
        plotted = d["ber"].where(d["ber"] > 0, np.nan)

        fig, ax = plt.subplots(figsize=(9.4, max(2.6, 0.34 * len(d) + 1.4)))
        ax.set_xscale("log")
        ax.set_xlim(x_lo, x_hi)
        # Позиции задаём числами: у нулевых мест марки нет, и категориальная ось
        # молча выбросила бы эти строки из графика.
        ypos = np.arange(len(d), dtype=float)
        # Высота < 1 оставляет зазор подложки между соседними марками.
        bars = ax.barh(ypos, plotted, height=0.62,
                       color=SERIES_COLORS[0], linewidth=0)
        ax.set_yticks(ypos)
        ax.set_yticklabels(d["label"])
        ax.set_ylim(-0.7, len(d) - 0.3)

        for y, raw in zip(ypos, d["ber"]):
            if raw > 0:
                ax.annotate(fmt_ber(raw), xy=(raw, y),
                            xytext=(5, 0), textcoords="offset points",
                            va="center", fontsize=7.5, color=INK_SECONDARY)
            else:
                ax.annotate("0 — ошибок не зафиксировано",
                            xy=(x_lo, y), xytext=(5, 0), textcoords="offset points",
                            va="center", ha="left", fontsize=7.5, color=STATUS_GOOD)
        for level, name, color in (
            (th.ber_pass_max, f"порог «эталонно чисто» {fmt_ber(th.ber_pass_max)}",
             STATUS_GOOD),
            (th.ber_warn_max, f"порог «нет запаса» {fmt_ber(th.ber_warn_max)}",
             STATUS_CRITICAL),
        ):
            ax.axvline(level, color=color, linewidth=1.0, zorder=1)
            ax.annotate(name, xy=(level, 1.005), xycoords=("data", "axes fraction"),
                        rotation=0, ha="center", va="bottom", fontsize=7,
                        color=INK_MUTED)

        ax.set_xlabel("Оценочный pre-FEC BER (логарифмическая шкала)")
        ax.set_title("Качество линка по всем посадочным местам", loc="left", pad=18)
        ax.grid(axis="y", visible=False)
        _despine(ax)
        fig.tight_layout()
        fig.savefig(out_path, bbox_inches="tight")
        plt.close(fig)
        return out_path
    except Exception as exc:  # noqa: BLE001
        LOG.error("Не удалось построить обзорный график BER: %s", exc)
        plt.close("all")
        return None


# ======================================================================================
#  11. ГЕНЕРАЦИЯ EXCEL
# ======================================================================================

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
THIN_SIDE = Side(style="thin", color="D9D9D9")
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


#: Человекочитаемые заголовки колонок для листов Excel.
COLUMN_TITLES: Dict[str, str] = {
    "serial": "Серийный номер",
    "serial_is_real": "S/N задан картой",
    "socket_id": "Идентификатор места",
    "test": "Круг",
    "switch": "Коммутатор",
    "side": "Сторона",
    "port": "Порт",
    "port_num": "№ порта",
    "source_file": "Файл лога",
    "tz": "Метка зоны",
    "fec_mode": "Режим FEC",
    "samples": "Снимков",
    "intervals_total": "Интервалов всего",
    "intervals_stable": "Интервалов стабильных",
    "intervals_nonstable": "Интервалов нештатных",
    "fec_uncorrected_stable": "FEC Uncorr. в стабильном прогоне",
    "fec_uncorrected_flap": "FEC Uncorr. при флапе/сбросе",
    "fec_uncorrected_linkdown": "FEC Uncorr. при потере линка",
    "carrier_transitions_stable": "Carrier trans. в стабильном прогоне",
    "fec_corrected_stable": "FEC Corr. в стабильном прогоне",
    "counter_baseline_corrected": "Стартовое значение FEC Corr.",
    "counter_baseline_uncorrected": "Стартовое значение FEC Uncorr.",
    "counter_zero_based": "Счётчики обнулены перед прогоном",
    "counter_resets": "Сбросов счётчика",
    "attribution": "Обстоятельства",
    "carrier_transitions": "Carrier transitions",
    "input_errors": "Input errors",
    "hs_link_crc_errors": "HS link CRC errors",
    "mtu_errors": "MTU errors",
    "resource_errors": "Resource errors",
    "fifo_errors": "FIFO errors",
    "bit_errors": "Bit errors",
    "errored_blocks": "Errored blocks",
    "crc_align_in": "CRC/Align вход",
    "crc_align_out": "CRC/Align выход",
    "attribution_code": "Категория",
    "stable": "Стабильный интервал",
    "ts_start": "Начало",
    "ts_end": "Окончание",
    "span_seconds": "Длительность захвата, с",
    "soak_seconds": "Полезная длительность, с",
    "soak_hours": "Полезная длительность, ч",
    "fec_corrected": "FEC Corrected (прирост)",
    "fec_uncorrected": "FEC Uncorrected (прирост)",
    "fec_corrected_rate_eps": "FEC Corrected, ошибок/с",
    "fec_uncorrected_rate_eps": "FEC Uncorrected, ошибок/с",
    "fec_corrected_rate_max_eps": "Пиковая скорость FEC Corr., ош./с",
    "fec_rate_cv": "Коэф. вариации потока FEC",
    "intervals_with_corrected": "Интервалов с ростом FEC Corr.",
    "ber_pre_fec_est": "Оценочный pre-FEC BER",
    "carrier_transitions_delta": "Прирост carrier transitions",
    "link_flaps": "Флапы линка",
    "link_down_samples": "Снимков с link Down",
    "defect_samples": "Снимков с Active defects",
    "hard_errors_total": "Жёстких ошибок всего",
    "verdict": "Вердикт",
    "final_verdict": "Итоговый вердикт",
    "rotation_conclusion": "Вывод по ротации",
    "reason_text": "Обоснование",
    "tests_participated": "Участие в кругах",
    "n_tests": "Кругов",
    "sockets": "Посадочные места",
    "verdict_per_test": "Вердикт по кругам",
    "fec_uncorrected_total": "FEC Uncorrected суммарно",
    "link_flaps_total": "Флапов суммарно",
    "ber_worst": "Худший pre-FEC BER",
    "soak_hours_total": "Наработка суммарно, ч",
    "dom_samples": "Снимков DOM",
    "lanes_seen": "Линий",
    "temp_min_c": "Температура мин., °C",
    "temp_max_c": "Температура макс., °C",
    "temp_mean_c": "Температура средн., °C",
    "temp_rise_c": "Прогрев Δt, °C",
    "tx_dbm_min": "TX мин., дБм",
    "tx_dbm_max": "TX макс., дБм",
    "tx_dbm_mean": "TX средн., дБм",
    "rx_dbm_min": "RX мин., дБм",
    "rx_dbm_max": "RX макс., дБм",
    "rx_dbm_mean": "RX средн., дБм",
    "bias_ma_min": "Bias мин., мА",
    "bias_ma_max": "Bias макс., мА",
    "bias_ma_mean": "Bias средн., мА",
    "tx_drift_db_per_day": "Дрейф TX, дБ/сут",
    "rx_drift_db_per_day": "Дрейф RX, дБ/сут",
    "bias_drift_ma_per_day": "Дрейф Bias, мА/сут",
    "tx_lane_imbalance_db": "Перекос TX между линиями, дБ",
    "rx_lane_imbalance_db": "Перекос RX между линиями, дБ",
    "tx_margin_low_db": "Запас TX до мин., дБ",
    "tx_margin_high_db": "Запас TX до макс., дБ",
    "rx_margin_low_db": "Запас RX до мин., дБ",
    "rx_margin_high_db": "Запас RX до макс., дБ",
    "bias_per_degc": "Чувствительность Bias, мА/°C",
    "link": "Линк",
    "link_verdict": "Вердикт линка",
    "a_socket": "Место A", "b_socket": "Место B",
    "a_serial": "S/N A", "b_serial": "S/N B",
    "a_verdict": "Вердикт A", "b_verdict": "Вердикт B",
    "a_fec_uncorrected": "FEC Uncorr. A", "b_fec_uncorrected": "FEC Uncorr. B",
    "a_ber": "BER A", "b_ber": "BER B",
    "a_flaps": "Флапы A", "b_flaps": "Флапы B",
    "link_dirty": "Линк проблемный",
    "ts_start_iv": "Начало интервала",
    "ts_end_iv": "Конец интервала",
    "duration_s": "Длительность, с",
    "usable": "Учтён в расчёте скоростей",
    "counter_reset": "Сброс счётчика",
    "link_disturbed": "Линк нарушен",
    "edge": "Краевое окно",
    "time_gap": "Пропуск мониторинга",
    "file": "Файл",
    "snapshot_index": "№ снимка",
    "timestamp": "Метка времени",
    "severity": "Уровень",
    "message": "Сообщение",
    "lane": "Линия",
    "temp_c": "Температура, °C",
    "bias_ma": "Bias, мА",
    "tx_dbm": "TX, дБм",
    "rx_dbm": "RX, дБм",
    "tx_mw": "TX, мВт",
    "rx_mw": "RX, мВт",
    "ts": "Метка времени",
    "snapshot": "№ снимка",
}

#: Колонки, к которым применяется цветовое кодирование вердикта.
VERDICT_COLUMNS = {"verdict", "final_verdict", "link_verdict", "a_verdict", "b_verdict"}

#: Формат чисел по колонкам.
NUMBER_FORMATS: Dict[str, str] = {
    "ber_pre_fec_est": "0.00E+00",
    "ber_worst": "0.00E+00",
    "a_ber": "0.00E+00",
    "b_ber": "0.00E+00",
    "fec_corrected": "# ##0",
    "fec_uncorrected": "# ##0",
    "fec_uncorrected_stable": "# ##0",
    "fec_uncorrected_flap": "# ##0",
    "fec_uncorrected_linkdown": "# ##0",
    "fec_corrected_stable": "# ##0",
    "carrier_transitions_stable": "# ##0",
    "counter_baseline_corrected": "# ##0",
    "counter_baseline_uncorrected": "# ##0",
    "fec_corrected_rate_eps": "# ##0.000",
    "fec_uncorrected_rate_eps": "0.000000",
    "fec_corrected_rate_max_eps": "# ##0.000",
    "fec_rate_cv": "0.00",
    "soak_hours": "0.0",
    "soak_hours_total": "0.0",
    "temp_min_c": "0.0", "temp_max_c": "0.0", "temp_mean_c": "0.0",
    "temp_rise_c": "+0.0;-0.0;0.0",
    "tx_dbm_min": "0.00", "tx_dbm_max": "0.00", "tx_dbm_mean": "0.00",
    "rx_dbm_min": "0.00", "rx_dbm_max": "0.00", "rx_dbm_mean": "0.00",
    "bias_ma_min": "0.000", "bias_ma_max": "0.000", "bias_ma_mean": "0.000",
    "tx_drift_db_per_day": "+0.000;-0.000;0.000",
    "rx_drift_db_per_day": "+0.000;-0.000;0.000",
    "bias_drift_ma_per_day": "+0.000;-0.000;0.000",
    "tx_lane_imbalance_db": "0.00", "rx_lane_imbalance_db": "0.00",
    "tx_margin_low_db": "0.00", "tx_margin_high_db": "0.00",
    "rx_margin_low_db": "0.00", "rx_margin_high_db": "0.00",
    "bias_per_degc": "+0.000;-0.000;0.000",
    "temp_c": "0.0", "bias_ma": "0.000",
    "tx_dbm": "0.00", "rx_dbm": "0.00", "tx_mw": "0.000", "rx_mw": "0.000",
    "ts": "DD.MM.YYYY HH:MM:SS",
    "ts_start": "DD.MM.YYYY HH:MM:SS",
    "ts_end": "DD.MM.YYYY HH:MM:SS",
}


def _style_sheet(ws: Worksheet, df: pd.DataFrame, freeze: str = "A2") -> None:
    """Единое оформление листа: шапка, ширины, форматы, автофильтр, закрепление."""
    if df.empty:
        return

    # --- шапка ----------------------------------------------------------------
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = COLUMN_TITLES.get(col_name, col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER

    ws.row_dimensions[1].height = 34
    ws.freeze_panes = freeze
    ws.auto_filter.ref = (f"A1:{get_column_letter(len(df.columns))}"
                          f"{len(df) + 1}")

    # --- ширины колонок --------------------------------------------------------
    for col_idx, col_name in enumerate(df.columns, start=1):
        title = COLUMN_TITLES.get(col_name, col_name)
        try:
            body_max = df[col_name].astype(str).str.len().max()
        except Exception:  # noqa: BLE001
            body_max = 12
        body_max = 12 if pd.isna(body_max) else int(body_max)
        # Заголовок переносится по словам, поэтому его вклад ограничиваем.
        width = min(max(len(title) * 0.62, body_max * 1.02, 10), 62)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # --- тело ------------------------------------------------------------------
    for col_idx, col_name in enumerate(df.columns, start=1):
        num_fmt = NUMBER_FORMATS.get(col_name)
        is_verdict = col_name in VERDICT_COLUMNS
        wrap = col_name in {"reason_text", "sockets", "verdict_per_test", "message",
                            "rotation_conclusion", "tests_participated"}
        for row_idx in range(2, len(df) + 2):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = CELL_BORDER
            if num_fmt:
                cell.number_format = num_fmt
            if wrap:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if is_verdict:
                verdict = cell.value
                fill = VERDICT_COLORS.get(verdict)
                if fill:
                    cell.fill = PatternFill("solid", fgColor=fill)
                    cell.font = Font(color=VERDICT_FONT_COLORS.get(verdict, "000000"),
                                     bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center",
                                           wrap_text=True)


def _prepare_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Готовит DataFrame к записи: убирает inf, приводит типы, вычисляет часы."""
    if df.empty:
        return df
    out = df.copy()
    out = out.replace([np.inf, -np.inf], np.nan)
    if "soak_seconds" in out.columns and "soak_hours" not in out.columns:
        out.insert(out.columns.get_loc("soak_seconds") + 1, "soak_hours",
                   out["soak_seconds"] / 3600.0)
    # Списки (причины) -> строки.
    for col in out.columns:
        if out[col].apply(lambda v: isinstance(v, (list, tuple, set))).any():
            out[col] = out[col].apply(
                lambda v: "; ".join(map(str, v)) if isinstance(v, (list, tuple, set)) else v)
    return out


def write_excel(
    path: Path,
    df_modules: pd.DataFrame,
    df_sockets: pd.DataFrame,
    df_dom_summary: pd.DataFrame,
    df_links: pd.DataFrame,
    df_intervals: pd.DataFrame,
    df_dom_raw: pd.DataFrame,
    issues: Sequence[ParseIssue],
    context: Dict[str, Any],
) -> Optional[Path]:
    """
    Записывает многолистовую книгу Excel.

    Листы:
        Общий итог            — вердикт по каждому модулю (S/N) с учётом ротации
        Сводка по местам      — метрики по каждому сокету (круг/свитч/порт)
        Сводка по DOM         — оптическая телеметрия, дрейф, перекосы
        Линки и ротация       — попарное сопоставление концов линка
        Временные ряды ошибок — поинтервальные дельты счётчиков
        Временные ряды DOM    — сырая телеметрия по линиям
        Методика и пороги     — какие критерии применялись
        Диагностика парсинга  — всё, что не удалось разобрать
    """
    try:
        with pd.ExcelWriter(path, engine="openpyxl", datetime_format="DD.MM.YYYY HH:MM:SS") as xl:
            sheets: List[Tuple[str, pd.DataFrame]] = []

            # --- 1. Общий итог -------------------------------------------------
            mod_cols = ["serial", "final_verdict", "rotation_conclusion",
                        "tests_participated", "verdict_per_test", "sockets",
                        "fec_uncorrected_total", "link_flaps_total", "ber_worst",
                        "soak_hours_total", "serial_is_real", "reason_text"]
            df_m = _prepare_for_excel(df_modules)
            df_m = df_m[[c for c in mod_cols if c in df_m.columns]]
            sheets.append(("Общий итог", df_m))

            # --- 2. Сводка по местам -------------------------------------------
            sock_cols = ["socket_id", "test", "switch", "port", "verdict",
                         "fec_uncorrected", "fec_uncorrected_stable",
                         "fec_uncorrected_flap", "fec_uncorrected_linkdown",
                         "link_flaps", "carrier_transitions_delta",
                         "carrier_transitions_stable", "link_down_samples",
                         "defect_samples", "hard_errors_total",
                         "fec_corrected", "fec_corrected_stable",
                         "fec_corrected_rate_eps", "fec_corrected_rate_max_eps",
                         "fec_rate_cv", "ber_pre_fec_est",
                         "samples", "intervals_stable", "intervals_nonstable",
                         "soak_seconds", "ts_start", "ts_end", "fec_mode",
                         "counter_zero_based", "counter_baseline_corrected",
                         "counter_baseline_uncorrected", "counter_resets",
                         "source_file", "reason_text"]
            df_s = _prepare_for_excel(df_sockets)
            keep = [c for c in sock_cols if c in df_s.columns]
            if "soak_hours" in df_s.columns:
                keep.insert(keep.index("soak_seconds") + 1, "soak_hours")
            sheets.append(("Сводка по местам", df_s[keep]))

            # --- 3. Сводка по DOM ----------------------------------------------
            dom_cols = ["socket_id", "test", "switch", "port", "dom_samples", "lanes_seen",
                        "temp_min_c", "temp_max_c", "temp_mean_c", "temp_rise_c",
                        "tx_dbm_min", "tx_dbm_max", "tx_dbm_mean",
                        "rx_dbm_min", "rx_dbm_max", "rx_dbm_mean",
                        "bias_ma_min", "bias_ma_max", "bias_ma_mean",
                        "tx_drift_db_per_day", "rx_drift_db_per_day",
                        "bias_drift_ma_per_day", "bias_per_degc",
                        "tx_lane_imbalance_db", "rx_lane_imbalance_db",
                        "tx_margin_low_db", "tx_margin_high_db",
                        "rx_margin_low_db", "rx_margin_high_db"]
            df_d = _prepare_for_excel(df_dom_summary)
            lane_cols = [c for c in df_d.columns if c.startswith("lane")]
            sheets.append(("Сводка по DOM",
                           df_d[[c for c in dom_cols if c in df_d.columns] + lane_cols]))

            # --- 4. Линки и ротация --------------------------------------------
            sheets.append(("Линки и ротация", _prepare_for_excel(df_links)))

            # --- 5. Временные ряды ошибок --------------------------------------
            iv_cols = ["test", "switch", "port", "ts_start", "ts_end", "duration_s",
                       "delta_fec_corrected", "delta_fec_uncorrected",
                       "delta_carrier_transitions", "delta_input_errors",
                       "delta_bit_errors", "delta_errored_blocks",
                       "delta_crc_align_in", "delta_hs_link_crc_errors",
                       "attribution", "stable", "counter_reset", "link_disturbed",
                       "edge", "time_gap"]
            df_i = _prepare_for_excel(df_intervals)
            df_i = df_i[[c for c in iv_cols if c in df_i.columns]]
            sheets.append(("Временные ряды ошибок", df_i))

            # --- 6. Временные ряды DOM -----------------------------------------
            dom_raw_cols = ["test", "switch", "port", "lane", "ts", "temp_c",
                            "bias_ma", "tx_dbm", "rx_dbm", "tx_mw", "rx_mw"]
            df_dr = _prepare_for_excel(df_dom_raw)
            df_dr = df_dr[[c for c in dom_raw_cols if c in df_dr.columns]]
            sheets.append(("Временные ряды DOM", df_dr))

            # --- 7. Реестр событий ошибок --------------------------------------
            # Каждый интервал, в котором вырос хоть один жёсткий счётчик, — с
            # временем и обстоятельствами. Именно здесь видно, что ошибка была,
            # даже если по скоростным метрикам место выглядит спокойным.
            ev_cols = ["test", "switch", "port", "ts_start", "ts_end", "attribution",
                       "fec_uncorrected", "fec_corrected", "carrier_transitions",
                       "counter_reset", "link_disturbed"] + HARD_ERROR_COLS
            df_ev = _prepare_for_excel(context.get("error_events", pd.DataFrame()))
            if df_ev.empty:
                df_ev = pd.DataFrame([{"test": "—", "switch": "—", "port": "—",
                                       "attribution": "Событий ошибок не зафиксировано"}])
            else:
                df_ev = df_ev[[c for c in ev_cols if c in df_ev.columns]]
            sheets.append(("Реестр событий ошибок", df_ev))

            # --- 8. Методика и пороги ------------------------------------------
            sheets.append(("Методика и пороги", context.get("methodology_df", pd.DataFrame())))

            # --- 9. Диагностика парсинга ---------------------------------------
            df_iss = pd.DataFrame([{
                "file": i.file, "snapshot_index": i.snapshot_index,
                "timestamp": i.timestamp, "severity": i.severity, "message": i.message,
            } for i in issues])
            if df_iss.empty:
                df_iss = pd.DataFrame([{
                    "file": "—", "snapshot_index": None, "timestamp": None,
                    "severity": "INFO",
                    "message": "Все снимки разобраны без ошибок.",
                }])
            sheets.append(("Диагностика парсинга", df_iss))

            for name, frame in sheets:
                if frame is None:
                    continue
                safe = frame if not frame.empty else pd.DataFrame({"Нет данных": []})
                safe.to_excel(xl, sheet_name=name[:31], index=False, header=False,
                              startrow=1)
                _style_sheet(xl.sheets[name[:31]], safe)

        LOG.info("Excel-отчёт сохранён: %s", path)
        return path
    except Exception as exc:  # noqa: BLE001
        LOG.error("Не удалось записать Excel-отчёт: %s", exc)
        LOG.debug("Трассировка:\n%s", traceback.format_exc())
        return None


# ======================================================================================
#  12. ГЕНЕРАЦИЯ PDF
# ======================================================================================

from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, Image, KeepTogether, NextPageTemplate, PageBreak,
    PageTemplate, Paragraph, Spacer, Table, TableStyle,
)

FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


def register_fonts() -> Tuple[str, str]:
    """
    Регистрирует TTF-шрифт с поддержкой кириллицы.

    Встроенные шрифты ReportLab (Helvetica и др.) кириллицу не содержат, поэтому
    берём DejaVu Sans — он поставляется вместе с matplotlib, то есть доступен
    везде, где работает этот скрипт. Предусмотрены и системные пути.
    """
    global FONT_REGULAR, FONT_BOLD
    candidates: List[Tuple[str, str]] = []
    try:
        mpl_ttf = Path(matplotlib.__file__).parent / "mpl-data" / "fonts" / "ttf"
        candidates.append((str(mpl_ttf / "DejaVuSans.ttf"),
                           str(mpl_ttf / "DejaVuSans-Bold.ttf")))
    except Exception:  # noqa: BLE001
        pass
    candidates += [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ]
    for regular, bold in candidates:
        if Path(regular).exists() and Path(bold).exists():
            try:
                pdfmetrics.registerFont(TTFont("ReportSans", regular))
                pdfmetrics.registerFont(TTFont("ReportSans-Bold", bold))
                FONT_REGULAR, FONT_BOLD = "ReportSans", "ReportSans-Bold"
                LOG.debug("Зарегистрирован шрифт: %s", regular)
                return FONT_REGULAR, FONT_BOLD
            except Exception as exc:  # noqa: BLE001
                LOG.debug("Шрифт %s не подошёл: %s", regular, exc)
    LOG.warning("TTF с кириллицей не найден — PDF будет в Helvetica, "
                "кириллица может отображаться некорректно.")
    return FONT_REGULAR, FONT_BOLD


def build_styles() -> Dict[str, ParagraphStyle]:
    """Набор стилей абзацев для отчёта."""
    base = getSampleStyleSheet()
    ink = rl_colors.HexColor("#0b0b0b")
    ink2 = rl_colors.HexColor("#52514e")
    muted = rl_colors.HexColor("#898781")
    return {
        "title": ParagraphStyle("title", parent=base["Title"], fontName=FONT_BOLD,
                                fontSize=26, leading=31, textColor=ink,
                                alignment=TA_LEFT, spaceAfter=4),
        "subtitle": ParagraphStyle("subtitle", parent=base["Normal"], fontName=FONT_REGULAR,
                                   fontSize=13, leading=18, textColor=ink2,
                                   alignment=TA_LEFT, spaceAfter=2),
        "kicker": ParagraphStyle("kicker", parent=base["Normal"], fontName=FONT_BOLD,
                                 fontSize=9, leading=12, textColor=muted,
                                 alignment=TA_LEFT, spaceAfter=6),
        "h1": ParagraphStyle("h1", parent=base["Heading1"], fontName=FONT_BOLD,
                             fontSize=15, leading=19, textColor=ink,
                             spaceBefore=2, spaceAfter=7),
        "h2": ParagraphStyle("h2", parent=base["Heading2"], fontName=FONT_BOLD,
                             fontSize=11.5, leading=15, textColor=ink,
                             spaceBefore=8, spaceAfter=4),
        "body": ParagraphStyle("body", parent=base["Normal"], fontName=FONT_REGULAR,
                               fontSize=9.3, leading=13.4, textColor=ink,
                               alignment=TA_JUSTIFY, spaceAfter=5),
        "small": ParagraphStyle("small", parent=base["Normal"], fontName=FONT_REGULAR,
                                fontSize=7.6, leading=10, textColor=ink2),
        "cell": ParagraphStyle("cell", parent=base["Normal"], fontName=FONT_REGULAR,
                               fontSize=7.4, leading=9.4, textColor=ink),
        "cell_b": ParagraphStyle("cell_b", parent=base["Normal"], fontName=FONT_BOLD,
                                 fontSize=7.4, leading=9.4, textColor=ink),
        "cell_c": ParagraphStyle("cell_c", parent=base["Normal"], fontName=FONT_BOLD,
                                 fontSize=7.4, leading=9.4, textColor=ink,
                                 alignment=TA_CENTER),
        "note": ParagraphStyle("note", parent=base["Normal"], fontName=FONT_REGULAR,
                               fontSize=8.2, leading=11.4,
                               textColor=rl_colors.HexColor("#9C0006")),
    }


def _para(text: Any, style: ParagraphStyle) -> Paragraph:
    """Безопасно оборачивает значение в Paragraph (экранирует спецсимволы XML)."""
    s = "" if text is None else str(text)
    s = (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    return Paragraph(s, style)


def _tint(hex_color: str, factor: float = 0.82) -> rl_colors.Color:
    """Осветляет цвет для использования в качестве фона ячейки таблицы."""
    c = rl_colors.HexColor(hex_color)
    return rl_colors.Color(c.red + (1 - c.red) * factor,
                           c.green + (1 - c.green) * factor,
                           c.blue + (1 - c.blue) * factor)


TABLE_BASE_STYLE = [
    ("FONTNAME", (0, 0), (-1, -1), FONT_REGULAR),
    ("FONTSIZE", (0, 0), (-1, -1), 7.4),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#e1e0d9")),
    ("TOPPADDING", (0, 0), (-1, -1), 3),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
]


def _header_style(ncols: int) -> List[Tuple]:
    return [
        ("BACKGROUND", (0, 0), (ncols - 1, 0), rl_colors.HexColor("#1F3864")),
        ("TEXTCOLOR", (0, 0), (ncols - 1, 0), rl_colors.white),
        ("FONTNAME", (0, 0), (ncols - 1, 0), FONT_BOLD),
        ("FONTSIZE", (0, 0), (ncols - 1, 0), 7.6),
        ("TOPPADDING", (0, 0), (ncols - 1, 0), 5),
        ("BOTTOMPADDING", (0, 0), (ncols - 1, 0), 5),
    ]


def build_executive_summary(
    df_modules: pd.DataFrame,
    df_sockets: pd.DataFrame,
    df_links: pd.DataFrame,
    context: Dict[str, Any],
) -> List[str]:
    """
    Формирует текст Executive Summary ПО ФАКТИЧЕСКИМ РЕЗУЛЬТАТАМ.

    Возвращает список абзацев. Ни одна формулировка не «зашита» заранее —
    все числа и выводы вычисляются из данных.
    """
    paras: List[str] = []
    profile: OpticProfile = context["profile"]
    map_active: bool = context["map_active"]

    n_sockets = len(df_sockets)
    n_modules = len(df_modules)
    n_tests = df_sockets["test"].nunique() if not df_sockets.empty else 0
    total_hours = (float(pd.to_numeric(df_sockets["soak_seconds"], errors="coerce")
                         .fillna(0).sum()) / 3600.0) if not df_sockets.empty else 0.0

    counts = (df_modules["final_verdict"].value_counts().to_dict()
              if not df_modules.empty else {})
    n_pass = counts.get(VERDICT_PASS, 0)
    n_warn = counts.get(VERDICT_WARN, 0)
    n_fail = counts.get(VERDICT_FAIL, 0)
    n_unk = counts.get(VERDICT_UNKNOWN, 0)

    # --- абзац 1: что и как тестировали ---------------------------------------
    paras.append(
        f"Проанализировано {n_sockets} посадочных мест в {n_tests} круг(ах) ротации; "
        f"суммарная наработка в стабильном режиме (линк Up, без флапов и сбросов "
        f"счётчика) — {total_hours:,.0f} часов".replace(",", " ") +
        f". Профиль оптики: {profile.name}. "
        f"Источники: {context.get('files_str', '—')}. "
        f"Интервал опроса — {context.get('poll_interval_str', '—')}."
    )

    # --- абзац 2: главный итог -------------------------------------------------
    unit = "модулей" if map_active else "посадочных мест"
    paras.append(
        f"<b>Итог:</b> из {n_modules} {unit} "
        f"<b>{n_pass}</b> рекомендованы в ЗИП, "
        f"<b>{n_warn}</b> признаны рабочими, но без запаса по бюджету трассы, "
        f"<b>{n_fail}</b> подлежат отбраковке" +
        (f", по {n_unk} данных недостаточно" if n_unk else "") + "."
    )

    # --- абзац 2а: неисправляемые FEC-ошибки, главный признак ------------------
    # Этот абзац стоит перед разбором вердиктов намеренно: неисправляемая
    # FEC-ошибка означает потерянный кадр и является первичным основанием
    # для отбраковки. Её нельзя прятать в примечания.
    if not df_sockets.empty and "fec_uncorrected" in df_sockets.columns:
        unc = pd.to_numeric(df_sockets["fec_uncorrected"], errors="coerce").fillna(0)
        unc_st = pd.to_numeric(df_sockets.get("fec_uncorrected_stable",
                                              pd.Series(0, index=df_sockets.index)),
                               errors="coerce").fillna(0)
        n_with = int((unc > 0).sum())
        n_stable = int((unc_st > 0).sum())
        if n_with:
            worst = df_sockets.loc[unc.idxmax()]
            listing = ", ".join(
                f"{r['switch']}:{port_short(r['port'])} (К{r['test']}) — "
                f"{fmt_count(u)}"
                for (_, r), u in sorted(
                    zip(df_sockets.iterrows(), unc), key=lambda z: -z[1])[:6] if u > 0)
            para = (f"<b>Неисправляемые FEC-ошибки зафиксированы на {n_with} из "
                    f"{n_sockets} посадочных мест</b>, суммарно "
                    f"{fmt_count(float(unc.sum()))}. Худшее: "
                    f"{worst['switch']}:{port_short(worst['port'])} "
                    f"(круг {worst['test']}) — {fmt_count(float(unc.max()))}. "
                    f"По местам: {listing}.")
            if n_stable:
                para += (f" <b>Из них на {n_stable} мест(ах) ошибки набраны при "
                         f"поднятом линке, без флапов и сбросов счётчика</b> — это "
                         f"свойство тракта, а не следствие монтажных работ.")
            else:
                para += (" Все они пришлись на интервалы с флапом линка, сбросом "
                         "счётчика или потерей сигнала — то есть на моменты "
                         "передёргивания линка. Обстоятельства каждого события "
                         "приведены в разделе «Реестр событий ошибок»; решение "
                         "о списании их на монтаж принимает инженер.")
            paras.append(para)
        else:
            paras.append("<b>Неисправляемых FEC-ошибок не зафиксировано ни на одном "
                         "посадочном месте за весь период наблюдения.</b>")

    # --- абзац 3: брак и его причины -------------------------------------------
    if n_fail:
        failed = df_modules[df_modules["final_verdict"] == VERDICT_FAIL]
        items = []
        for _, r in failed.head(8).iterrows():
            first_reason = str(r.get("reason_text", "")).split("||")[0].strip()
            first_reason = re.sub(r"^\[[^\]]+\]\s*", "", first_reason)
            first_reason = first_reason.split("|")[0].strip()
            items.append(f"<b>{r['serial']}</b> — {first_reason}")
        paras.append("<b>Отбраковка.</b> " + "; ".join(items) + ".")
    else:
        paras.append("<b>Отбраковка.</b> Ни одно посадочное место не набрало "
                     "браковочных признаков: неисправляемых FEC-ошибок, флапов линка, "
                     "выходов DOM за границы даташита и деградации мощности "
                     "в пределах анализируемых окон не зафиксировано.")

    # --- абзац 4: предупреждения ------------------------------------------------
    if n_warn:
        warned = df_modules[df_modules["final_verdict"] == VERDICT_WARN]
        names = ", ".join(str(s) for s in warned["serial"].head(10))
        worst_ber = pd.to_numeric(warned.get("ber_worst", pd.Series(dtype=float)),
                                  errors="coerce").max()
        paras.append(
            f"<b>Не в ЗИП.</b> {names} — линк держится устойчиво и неисправляемых "
            f"ошибок нет, однако поток исправленных FEC-ошибок значим "
            f"(худший оценочный pre-FEC BER {fmt_ber(float(worst_ber) if pd.notna(worst_ber) else None)}). "
            f"Такие модули пригодны для некритичных сегментов с коротким "
            f"бюджетом трассы, но не как холодный резерв магистрали."
        )

    # --- абзац 5: локализация по ротации ---------------------------------------
    if map_active:
        confirmed = df_modules[df_modules["rotation_conclusion"] == ROT_CONFIRMED_FAIL]
        link_fault = df_modules[df_modules["rotation_conclusion"] == ROT_LINK_FAULT]
        retest = df_modules[df_modules["rotation_conclusion"] == ROT_NEEDS_RETEST]
        bits = []
        if len(confirmed):
            bits.append(f"{len(confirmed)} модул(ей) дали ошибки в двух кругах "
                        f"с разными партнёрами — вина модуля доказана "
                        f"({', '.join(map(str, confirmed['serial'].head(6)))})")
        if len(link_fault):
            bits.append(f"{len(link_fault)} модул(ей) сыпали в одном круге и были "
                        f"чисты в другом — виновата трасса (патч-корд/загрязнение "
                        f"коннектора), модули годны после перечистки "
                        f"({', '.join(map(str, link_fault['serial'].head(6)))})")
        if len(retest):
            bits.append(f"{len(retest)} модул(ей) требуют повторного прогона: "
                        f"признаки есть, но второго круга для перекрёстной проверки нет")
        paras.append("<b>Локализация по методике двух кругов.</b> " +
                     ("; ".join(bits) + "." if bits else
                      "Расхождений между кругами не выявлено — результаты воспроизводимы."))
    else:
        paras.append(
            "<b>Локализация по методике двух кругов не выполнена.</b> В выводе команд "
            "<i>show interfaces ... extensive</i> и <i>show interfaces diagnostics optics</i> "
            "серийные номера трансиверов отсутствуют, а карта ротации "
            "(--map modules_map.csv) не задана. Поэтому каждое посадочное место "
            "оценено независимо, и разделить «виноват модуль» и «виновата трасса» "
            "невозможно. Заполните карту ротации и перезапустите анализ — "
            "кросс-раундовые выводы будут построены автоматически."
        )

    # --- абзац 6: методические оговорки ----------------------------------------
    notes = []
    resets = (int(context.get("counter_resets", 0)))
    if resets:
        notes.append(
            f"Обнаружено {resets} сброс(ов) кумулятивных счётчиков Junos "
            f"(clear interfaces statistics либо пересоздание интерфейса). "
            f"Поэтому все метрики построены на поинтервальных дельтах: разница "
            f"«первое значение минус последнее» дала бы заведомо неверный результат."
        )
    not_zero = (int((~df_sockets["counter_zero_based"].fillna(True)).sum())
                if "counter_zero_based" in df_sockets.columns else 0)
    if not_zero:
        notes.append(
            f"У {not_zero} посадочных мест счётчики не были обнулены перед началом "
            f"прогона — стартовое значение унаследовано от предыдущего круга. "
            f"В метриках учитывается прирост за захват, а не абсолютное значение "
            f"счётчика; абсолютные величины приведены отдельными колонками."
        )
    if context.get("tz_conflict"):
        notes.append(
            f"Стороны линка помечены разными временными зонами "
            f"({context['tz_conflict']}), при этом стенные часы идут синхронно — "
            f"метка зоны на одном из коммутаторов выставлена неверно. Корреляция "
            f"концов выполнена по стенному времени; рекомендуется привести NTP "
            f"и зону к единому виду перед следующим прогоном."
        )
    if notes:
        paras.append("<b>Методические оговорки.</b> " + " ".join(notes))

    return paras


def make_verdict_table(
    df_modules: pd.DataFrame, styles: Dict[str, ParagraphStyle], width: float,
) -> Table:
    """Сводная таблица вердиктов с цветовым кодированием (цвет + текстовая метка)."""
    header = ["Модуль / место", "Вердикт", "Вывод по ротации", "Круги",
              "FEC Uncorr.", "Флапы", "Худший BER", "Наработка, ч"]
    rows: List[List[Any]] = [[_para(h, styles["cell_c"]) for h in header]]
    style_cmds: List[Tuple] = list(TABLE_BASE_STYLE) + _header_style(len(header))

    for i, (_, r) in enumerate(df_modules.iterrows(), start=1):
        verdict = r.get("final_verdict", VERDICT_UNKNOWN)
        rows.append([
            _para(r.get("serial"), styles["cell_b"]),
            _para(verdict, styles["cell_c"]),
            _para(r.get("rotation_conclusion"), styles["cell"]),
            _para(r.get("tests_participated"), styles["cell"]),
            _para(fmt_count(r.get("fec_uncorrected_total")), styles["cell_c"]),
            _para(fmt_count(r.get("link_flaps_total")), styles["cell_c"]),
            _para(fmt_ber(r.get("ber_worst")), styles["cell_c"]),
            _para(fmt_num(r.get("soak_hours_total"), 1), styles["cell_c"]),
        ])
        status = STATUS_BY_VERDICT.get(verdict, "#898781")
        style_cmds.append(("BACKGROUND", (1, i), (1, i), _tint(status, 0.72)))
        # Левая цветная «метка» строки — дублирует статус, но смысл несёт текст.
        style_cmds.append(("LINEBEFORE", (0, i), (0, i), 3, rl_colors.HexColor(status)))

    col_widths = [width * w for w in (0.20, 0.135, 0.245, 0.10, 0.075, 0.055, 0.10, 0.085)]
    table = Table(rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle(style_cmds))
    return table


def make_socket_table(
    df_sockets: pd.DataFrame, styles: Dict[str, ParagraphStyle], width: float,
) -> Table:
    """Подробная таблица метрик по каждому посадочному месту."""
    header = ["Круг", "Свитч", "Порт", "Вердикт", "FEC Uncorr.\nвсего",
              "из них в\nстабильном", "Флапы", "FEC Corr.", "Скорость, ош./с",
              "pre-FEC BER", "Стаб., ч", "Стаб. инт."]
    rows: List[List[Any]] = [[_para(h, styles["cell_c"]) for h in header]]
    style_cmds: List[Tuple] = list(TABLE_BASE_STYLE) + _header_style(len(header))

    for i, (_, r) in enumerate(df_sockets.iterrows(), start=1):
        verdict = r.get("verdict", VERDICT_UNKNOWN)
        rows.append([
            _para(f"К{r.get('test')}", styles["cell_c"]),
            _para(r.get("switch"), styles["cell_c"]),
            _para(port_short(r.get("port")), styles["cell_c"]),
            _para(verdict, styles["cell_c"]),
            _para(fmt_count(r.get("fec_uncorrected")),
                  styles["cell_b"] if float(r.get("fec_uncorrected") or 0) > 0
                  else styles["cell_c"]),
            _para(fmt_count(r.get("fec_uncorrected_stable")),
                  styles["cell_b"] if float(r.get("fec_uncorrected_stable") or 0) > 0
                  else styles["cell_c"]),
            _para(fmt_count(r.get("link_flaps")), styles["cell_c"]),
            _para(fmt_count(r.get("fec_corrected")), styles["cell_c"]),
            _para(fmt_num(r.get("fec_corrected_rate_eps"), 1), styles["cell_c"]),
            _para(fmt_ber(r.get("ber_pre_fec_est")), styles["cell_c"]),
            _para(fmt_num(float(r.get("soak_seconds") or 0) / 3600.0, 1), styles["cell_c"]),
            _para(f"{int(r.get('intervals_stable') or 0)}/{int(r.get('intervals_total') or 0)}",
                  styles["cell_c"]),
        ])
        status = STATUS_BY_VERDICT.get(verdict, "#898781")
        style_cmds.append(("BACKGROUND", (3, i), (3, i), _tint(status, 0.72)))
        # Ненулевые неисправляемые ошибки подсвечиваем всегда: это первичный
        # признак браковки, и он не должен теряться среди прочих колонок.
        if float(r.get("fec_uncorrected") or 0) > 0:
            style_cmds.append(("BACKGROUND", (4, i), (4, i),
                               _tint(STATUS_WARNING, 0.70)))
        if float(r.get("fec_uncorrected_stable") or 0) > 0:
            style_cmds.append(("BACKGROUND", (5, i), (5, i),
                               _tint(STATUS_CRITICAL, 0.70)))

    col_widths = [width * w for w in
                  (0.05, 0.065, 0.05, 0.125, 0.085, 0.085, 0.055, 0.095, 0.09,
                   0.095, 0.06, 0.065)]
    table = Table(rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle(style_cmds))
    return table


def make_events_table(
    df_events: pd.DataFrame, styles: Dict[str, ParagraphStyle], width: float,
    limit: int = 40,
) -> Table:
    """
    Реестр событий ошибок: когда, где, сколько и при каких обстоятельствах.

    Это ключевая таблица отчёта. Она отвечает на вопрос, который сводные
    метрики скрывают: неисправляемая FEC-ошибка была зафиксирована в стабильном
    прогоне или в момент, когда линк передёргивали.
    """
    header = ["Круг", "Свитч", "Порт", "Время события", "Обстоятельства",
              "FEC Uncorr.", "FEC Corr.", "Carrier trans."]
    rows: List[List[Any]] = [[_para(h, styles["cell_c"]) for h in header]]
    style_cmds: List[Tuple] = list(TABLE_BASE_STYLE) + _header_style(len(header))

    shown = df_events.head(limit)
    for i, (_, r) in enumerate(shown.iterrows(), start=1):
        unc = float(r.get("fec_uncorrected") or 0)
        code = r.get("attribution_code", "")
        rows.append([
            _para(f"К{r.get('test')}", styles["cell_c"]),
            _para(r.get("switch"), styles["cell_c"]),
            _para(port_short(r.get("port")), styles["cell_c"]),
            _para(f"{r['ts_end']:%d.%m %H:%M}" if pd.notna(r.get("ts_end")) else "—",
                  styles["cell_c"]),
            _para(r.get("attribution"), styles["cell"]),
            _para(fmt_count(unc), styles["cell_b"] if unc > 0 else styles["cell_c"]),
            _para(fmt_count(r.get("fec_corrected")), styles["cell_c"]),
            _para(fmt_count(r.get("carrier_transitions")), styles["cell_c"]),
        ])
        # Красным выделяется только самый тяжёлый случай: неисправляемые ошибки
        # в стабильном прогоне. Цвет сопровождается текстом в колонке
        # «Обстоятельства», то есть смысл не несётся одним лишь цветом.
        if unc > 0:
            tone = STATUS_CRITICAL if code == "STABLE" else STATUS_WARNING
            style_cmds.append(("BACKGROUND", (5, i), (5, i), _tint(tone, 0.70)))
            style_cmds.append(("LINEBEFORE", (0, i), (0, i), 3,
                               rl_colors.HexColor(tone)))

    col_widths = [width * w for w in
                  (0.055, 0.075, 0.055, 0.115, 0.30, 0.14, 0.13, 0.13)]
    table = Table(rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle(style_cmds))
    return table


def make_link_table(
    df_links: pd.DataFrame, styles: Dict[str, ParagraphStyle], width: float,
) -> Table:
    """Таблица линков: оба конца рядом — так виден «виновник» пары."""
    header = ["Круг", "Линк", "Вердикт линка", "Конец A", "Вердикт A",
              "Конец B", "Вердикт B"]
    rows: List[List[Any]] = [[_para(h, styles["cell_c"]) for h in header]]
    style_cmds: List[Tuple] = list(TABLE_BASE_STYLE) + _header_style(len(header))

    for i, (_, r) in enumerate(df_links.iterrows(), start=1):
        lv = r.get("link_verdict", VERDICT_UNKNOWN)
        rows.append([
            _para(f"К{r.get('test')}", styles["cell_c"]),
            _para(r.get("link"), styles["cell_b"]),
            _para(lv, styles["cell_c"]),
            _para(r.get("a_serial") or r.get("a_socket"), styles["cell"]),
            _para(r.get("a_verdict"), styles["cell_c"]),
            _para(r.get("b_serial") or r.get("b_socket"), styles["cell"]),
            _para(r.get("b_verdict"), styles["cell_c"]),
        ])
        for col, key in ((2, "link_verdict"), (4, "a_verdict"), (6, "b_verdict")):
            status = STATUS_BY_VERDICT.get(r.get(key), "#898781")
            style_cmds.append(("BACKGROUND", (col, i), (col, i), _tint(status, 0.75)))

    col_widths = [width * w for w in (0.06, 0.16, 0.145, 0.165, 0.145, 0.165, 0.145)]
    table = Table(rows, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle(style_cmds))
    return table


class ReportDocTemplate(BaseDocTemplate):
    """Шаблон документа с колонтитулами и сквозной нумерацией страниц."""

    def __init__(self, filename: str, **kwargs: Any) -> None:
        self.report_title = kwargs.pop("report_title", "Отчёт")
        super().__init__(filename, **kwargs)
        frame = Frame(self.leftMargin, self.bottomMargin,
                      self.width, self.height, id="main")
        self.addPageTemplates([
            PageTemplate(id="cover", frames=[frame], onPage=self._on_cover),
            PageTemplate(id="body", frames=[frame], onPage=self._on_body),
        ])

    def _on_cover(self, canvas, doc) -> None:  # noqa: ANN001
        canvas.saveState()
        canvas.setFillColor(rl_colors.HexColor("#1F3864"))
        canvas.rect(0, doc.pagesize[1] - 14 * mm, doc.pagesize[0], 14 * mm,
                    stroke=0, fill=1)
        canvas.restoreState()

    def _on_body(self, canvas, doc) -> None:  # noqa: ANN001
        canvas.saveState()
        canvas.setFont(FONT_REGULAR, 7.4)
        canvas.setFillColor(rl_colors.HexColor("#898781"))
        # Верхний колонтитул.
        canvas.drawString(doc.leftMargin, doc.pagesize[1] - 11 * mm, self.report_title)
        canvas.setStrokeColor(rl_colors.HexColor("#e1e0d9"))
        canvas.setLineWidth(0.5)
        canvas.line(doc.leftMargin, doc.pagesize[1] - 12.5 * mm,
                    doc.pagesize[0] - doc.rightMargin, doc.pagesize[1] - 12.5 * mm)
        # Нижний колонтитул.
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, 10 * mm,
                               f"стр. {doc.page}")
        canvas.drawString(doc.leftMargin, 10 * mm,
                          "Автоматический анализ soak-теста QSFP28")
        canvas.restoreState()


def build_pdf(
    path: Path,
    df_modules: pd.DataFrame,
    df_sockets: pd.DataFrame,
    df_dom_summary: pd.DataFrame,
    df_links: pd.DataFrame,
    df_intervals: pd.DataFrame,
    df_dom_raw: pd.DataFrame,
    issues: Sequence[ParseIssue],
    context: Dict[str, Any],
    charts_dir: Path,
) -> Optional[Path]:
    """Собирает презентабельный аналитический PDF-отчёт."""
    try:
        register_fonts()
        styles = build_styles()
        # Стиль таблиц опирается на зарегистрированные шрифты — обновляем.
        for i, cmd in enumerate(TABLE_BASE_STYLE):
            if cmd[0] == "FONTNAME":
                TABLE_BASE_STYLE[i] = ("FONTNAME", (0, 0), (-1, -1), FONT_REGULAR)

        page = landscape(A4)
        doc = ReportDocTemplate(
            str(path), pagesize=page,
            leftMargin=14 * mm, rightMargin=14 * mm,
            topMargin=16 * mm, bottomMargin=14 * mm,
            title="Отчёт по приёмочному soak-тесту QSFP28",
            author="Network Automation",
            report_title=context.get("report_title", "Soak-тест QSFP28"),
        )
        avail = doc.width
        story: List[Any] = []

        # ------------------------------------------------------------- титул ---
        story.append(Spacer(1, 26 * mm))
        story.append(_para("ПРИЁМОЧНОЕ ТЕСТИРОВАНИЕ ОПТИЧЕСКИХ ТРАНСИВЕРОВ",
                           styles["kicker"]))
        # Здесь нужен управляемый перенос строки, поэтому разметка передаётся
        # в Paragraph напрямую, минуя экранирование _para().
        story.append(Paragraph("Отчёт по back-to-back soak-тесту<br/>"
                               "100G QSFP28 (Upnet)", styles["title"]))
        story.append(Spacer(1, 4 * mm))
        story.append(_para(
            f"Коммутаторы Juniper QFX5110-48S-4C · методика ротации в "
            f"{context.get('n_tests', 0)} круг(а) · профиль оптики "
            f"{context['profile'].name}", styles["subtitle"]))
        story.append(Spacer(1, 10 * mm))

        counts = (df_modules["final_verdict"].value_counts().to_dict()
                  if not df_modules.empty else {})
        cover_rows = [
            ["Дата формирования отчёта", datetime.now().strftime("%d.%m.%Y %H:%M")],
            ["Период тестирования", context.get("period_str", "—")],
            ["Файлы логов", context.get("files_str", "—")],
            ["Интервал опроса", context.get("poll_interval_str", "—")],
            ["Посадочных мест / модулей", f"{len(df_sockets)} / {len(df_modules)}"],
            ["Полезная наработка суммарно", context.get("total_soak_str", "—")],
            ["Привязка к серийным номерам",
             "по карте ротации" if context["map_active"]
             else "НЕ ВЫПОЛНЕНА — S/N в логах отсутствуют"],
            ["Результат",
             f"в ЗИП: {counts.get(VERDICT_PASS, 0)}   ·   "
             f"не в ЗИП: {counts.get(VERDICT_WARN, 0)}   ·   "
             f"брак: {counts.get(VERDICT_FAIL, 0)}"],
        ]
        t = Table([[_para(a, styles["cell_b"]), _para(b, styles["cell"])]
                   for a, b in cover_rows],
                  colWidths=[avail * 0.26, avail * 0.54], hAlign="LEFT")
        t.setStyle(TableStyle(list(TABLE_BASE_STYLE) + [
            ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#e1e0d9")),
            ("BACKGROUND", (0, 0), (0, -1), rl_colors.HexColor("#f4f4f1")),
            ("FONTSIZE", (0, 0), (-1, -1), 8.4),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t)
        # Со второй страницы — шаблон с колонтитулами.
        story.append(NextPageTemplate("body"))
        story.append(PageBreak())

        # --------------------------------------------- Executive Summary ------
        story.append(_para("EXECUTIVE SUMMARY", styles["kicker"]))
        story.append(_para("Краткое заключение", styles["h1"]))
        for p in build_executive_summary(df_modules, df_sockets, df_links, context):
            story.append(Paragraph(p, styles["body"]))

        story.append(Spacer(1, 4 * mm))
        story.append(_para("Сводные вердикты", styles["h2"]))
        if not df_modules.empty:
            story.append(make_verdict_table(df_modules, styles, avail))
            story.append(Spacer(1, 2.5 * mm))
            story.append(_para(
                "Цвет строки дублирует текстовую метку вердикта и не является "
                "единственным носителем смысла: «В ЗИП» — зелёный, «Рабочий, не в ЗИП» — "
                "жёлтый, «БРАК / Утилизация» — красный.", styles["small"]))
        story.append(PageBreak())

        # ------------------------------------------------ обзор по местам -----
        story.append(_para("ДЕТАЛИЗАЦИЯ", styles["kicker"]))
        story.append(_para("Метрики по посадочным местам", styles["h1"]))
        if not df_sockets.empty:
            story.append(make_socket_table(df_sockets, styles, avail))
        story.append(Spacer(1, 4 * mm))
        ber_chart = charts_dir / "overview_ber.png"
        if ber_chart.exists():
            story.append(_fit_image(ber_chart, avail, doc.height - 14 * mm))
        story.append(PageBreak())

        # ----------------------------------------------------- линки ----------
        if not df_links.empty:
            story.append(_para("ЛОКАЛИЗАЦИЯ", styles["kicker"]))
            story.append(_para("Линки и результат ротации", styles["h1"]))
            story.append(_para(
                "Таблица сводит оба конца каждого линка. Если «грязным» оказался "
                "только один конец — подозрение падает на модуль этого конца; если "
                "оба — на трассу между ними. Окончательный вывод даёт сопоставление "
                "с результатами второго круга (колонка «Вывод по ротации» в сводке).",
                styles["body"]))
            story.append(make_link_table(df_links, styles, avail))
            story.append(PageBreak())

        # ------------------------------------------- реестр событий ошибок -----
        df_events = context.get("error_events", pd.DataFrame())
        story.append(_para("СОБЫТИЯ", styles["kicker"]))
        story.append(_para("Реестр событий ошибок", styles["h1"]))
        if df_events is None or df_events.empty:
            story.append(_para(
                "За весь период наблюдения не зафиксировано ни одного прироста "
                "неисправляемых FEC-ошибок, carrier transitions или input/CRC/bit "
                "errors.", styles["body"]))
        else:
            n_unc = int((pd.to_numeric(df_events["fec_uncorrected"],
                                       errors="coerce").fillna(0) > 0).sum())
            n_stable = int(((pd.to_numeric(df_events["fec_uncorrected"],
                                           errors="coerce").fillna(0) > 0)
                            & (df_events["attribution_code"] == "STABLE")).sum())
            story.append(_para(
                f"Зафиксировано {len(df_events)} интервал(ов) с приростом жёстких "
                f"счётчиков, из них {n_unc} с неисправляемыми FEC-ошибками "
                f"({n_stable} — в стабильном прогоне, то есть при поднятом линке, "
                f"без флапов и сбросов счётчика). Колонка «Обстоятельства» "
                f"определяется по состоянию линка и счётчиков в самом логе, а не "
                f"по предположениям о действиях оператора. Ни одно из этих событий "
                f"не исключается из вердикта: неисправляемая FEC-ошибка означает "
                f"потерянный кадр.", styles["body"]))
            story.append(make_events_table(df_events, styles, avail))
            if len(df_events) > 40:
                story.append(Spacer(1, 2 * mm))
                story.append(_para(
                    f"Показаны первые 40 событий из {len(df_events)}. "
                    f"Полный реестр — на листе «Реестр событий ошибок» в Excel.",
                    styles["small"]))
        story.append(PageBreak())

        # ------------------------------------------- постраничная детализация --
        story.append(_para("ГРАФИКИ", styles["kicker"]))
        story.append(_para("Динамика DOM и ошибок по каждому посадочному месту",
                           styles["h1"]))
        story.append(_para(
            "Для каждого места приведены: прогрев (температура и ток смещения по "
            "линиям), оптическая мощность TX/RX по каждой из четырёх линий QSFP28 "
            "с границами даташита, и временной ряд FEC-ошибок. Серой заливкой "
            "отмечены нештатные интервалы (флап линка, сброс счётчика, потеря "
            "сигнала): данные в них выводятся и учитываются в вердикте — заливка "
            "указывает на обстоятельства, а не исключает значение. Каждое событие "
            "неисправляемых ошибок помечено маркой: круг — в стабильном прогоне, "
            "треугольник — в нештатном интервале.",
            styles["body"]))
        story.append(Spacer(1, 2 * mm))

        for _, r in df_sockets.iterrows():
            sid = r["socket_id"]
            slug = re.sub(r"[^A-Za-z0-9]+", "_", sid).strip("_")
            serial = r.get("serial", sid)
            verdict = r.get("verdict", VERDICT_UNKNOWN)

            story.append(PageBreak())
            heading = (f"Круг {r['test']} · {r['switch']} · порт {r['port']}"
                       + (f" · S/N {serial}" if context["map_active"] else ""))
            story.append(_para(heading, styles["h1"]))

            status = STATUS_BY_VERDICT.get(verdict, "#898781")
            badge = Table([[_para(f"Вердикт: {verdict}", styles["cell_b"]),
                            _para(r.get("reason_text", ""), styles["cell"])]],
                          colWidths=[avail * 0.20, avail * 0.80], hAlign="LEFT")
            badge.setStyle(TableStyle(list(TABLE_BASE_STYLE) + [
                ("BACKGROUND", (0, 0), (0, 0), _tint(status, 0.72)),
                ("LINEBEFORE", (0, 0), (0, 0), 3, rl_colors.HexColor(status)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))
            story.append(badge)
            story.append(Spacer(1, 3 * mm))

            # Ширина подобрана так, чтобы на странице умещалось два графика:
            # при полной ширине каждый занимал бы целый лист, и отчёт разбухал
            # вдвое при той же информативности.
            img_w = avail * 0.62
            for suffix in ("thermal", "power", "fec"):
                img = charts_dir / f"{slug}_{suffix}.png"
                if img.exists():
                    story.append(_fit_image(img, img_w, doc.height - 6 * mm))
                    story.append(Spacer(1, 2.0 * mm))

        # ---------------------------------------------- методика приложением ---
        story.append(PageBreak())
        story.append(_para("ПРИЛОЖЕНИЕ", styles["kicker"]))
        story.append(_para("Методика и пороги браковки", styles["h1"]))
        meth = context.get("methodology_df", pd.DataFrame())
        if not meth.empty:
            rows = [[_para("Параметр", styles["cell_c"]),
                     _para("Значение", styles["cell_c"]),
                     _para("Комментарий", styles["cell_c"])]]
            for _, m in meth.iterrows():
                rows.append([_para(m.iloc[0], styles["cell_b"]),
                             _para(m.iloc[1], styles["cell"]),
                             _para(m.iloc[2] if len(m) > 2 else "", styles["cell"])])
            t = Table(rows, colWidths=[avail * 0.28, avail * 0.18, avail * 0.54],
                      repeatRows=1, hAlign="LEFT")
            t.setStyle(TableStyle(list(TABLE_BASE_STYLE) + _header_style(3)))
            story.append(t)

        # --------------------------------------------- диагностика парсинга ----
        err_issues = [i for i in issues if i.severity in ("ERROR", "WARNING")]
        story.append(Spacer(1, 4 * mm))
        story.append(_para("Диагностика разбора логов", styles["h2"]))
        if err_issues:
            story.append(_para(
                f"При разборе зафиксировано {len(err_issues)} замечани(й). "
                f"Полный список — на листе «Диагностика парсинга» в Excel-книге.",
                styles["note"]))
            rows = [[_para(h, styles["cell_c"]) for h in
                     ("Файл", "№ снимка", "Уровень", "Сообщение")]]
            for i in err_issues[:25]:
                rows.append([_para(i.file, styles["cell"]),
                             _para(i.snapshot_index, styles["cell_c"]),
                             _para(i.severity, styles["cell_c"]),
                             _para(i.message, styles["cell"])])
            t = Table(rows, colWidths=[avail * 0.18, avail * 0.08,
                                       avail * 0.09, avail * 0.65],
                      repeatRows=1, hAlign="LEFT")
            t.setStyle(TableStyle(list(TABLE_BASE_STYLE) + _header_style(4)))
            story.append(t)
        else:
            story.append(_para("Все снимки во всех файлах разобраны без ошибок.",
                               styles["body"]))

        # BaseDocTemplate управляет колонтитулами через PageTemplate,
        # поэтому onFirstPage/onLaterPages здесь не применяются.
        doc.build(story)
        LOG.info("PDF-отчёт сохранён: %s", path)
        return path
    except Exception as exc:  # noqa: BLE001
        LOG.error("Не удалось собрать PDF-отчёт: %s", exc)
        LOG.debug("Трассировка:\n%s", traceback.format_exc())
        return None


def _fit_image(path: Path, max_w: float, max_h: float) -> Image:
    """
    Вставляет картинку, вписывая её в доступную область.

    Без этого высокий график (например, обзорная панель на много строк) при
    полной ширине превышает высоту фрейма, и ReportLab отклоняет весь документ.
    """
    ratio = _img_ratio(path)
    w = max_w
    h = w * ratio
    if h > max_h:
        h = max_h
        w = h / ratio if ratio else max_w
    return Image(str(path), width=w, height=h)


def _img_ratio(path: Path) -> float:
    """Отношение высоты к ширине PNG — чтобы вставлять картинку без искажений."""
    try:
        from reportlab.lib.utils import ImageReader
        w, h = ImageReader(str(path)).getSize()
        return h / w if w else 0.5
    except Exception:  # noqa: BLE001
        return 0.5


# ======================================================================================
#  13. ОРКЕСТРАЦИЯ
# ======================================================================================

def build_methodology_table(
    profile: OpticProfile, th: Thresholds, context: Dict[str, Any],
) -> pd.DataFrame:
    """Таблица «что и по каким порогам проверялось» — для Excel и приложения PDF."""
    rows = [
        ("Профиль оптики", profile.name,
         "Границы даташита, по которым проверялись TX/RX/Bias/температура."),
        ("TX power, дБм", f"{profile.tx_dbm_min} … {profile.tx_dbm_max}",
         "Выход за границы — безусловная отбраковка."),
        ("RX power, дБм", f"{profile.rx_dbm_min} … {profile.rx_dbm_max}",
         "Выход за границы — безусловная отбраковка."),
        ("Bias current, мА", f"{profile.bias_ma_min} … {profile.bias_ma_max}",
         f"Типовой максимум {profile.bias_ma_typ_max} мА; превышение — предупреждение."),
        ("Температура, °C", f"{profile.temp_c_min} … {profile.temp_c_max}",
         "Диапазон коммерческого исполнения."),
        ("FEC Uncorrected", f"> {th.uncorrected_fail}",
         "Любая неисправляемая FEC-ошибка в пригодном интервале — брак."),
        ("pre-FEC BER «эталон»", fmt_ber(th.ber_pass_max),
         "Ниже порога — линк считается эталонно чистым (в ЗИП)."),
        ("pre-FEC BER «без запаса»", fmt_ber(th.ber_warn_max),
         f"Выше — брак. Между порогами — «рабочий, не в ЗИП». "
         f"Ориентир: бюджет FEC91 ≈ {fmt_ber(FEC91_PREFEC_BER_BUDGET)}."),
        ("Скорость линии для BER", f"{LINE_RATE_BPS/1e9:.3f} Гбит/с",
         "Пересчёт «FEC corrected ошибок/с» в оценочный pre-FEC BER."),
        ("Коэф. вариации потока FEC", f"> {th.fec_rate_cv_warn}",
         "Признак нестабильного, всплескового потока ошибок."),
        ("Флапы линка", f"> {th.link_flaps_fail}",
         "Считаются по приросту carrier transitions вне краевых окон (2 перехода = 1 флап)."),
        ("Input / CRC / Bit errors", f"> {th.input_errors_fail}",
         "Любой ненулевой прирост — брак."),
        ("Дрейф TX/RX, дБ/сут", f"warn {th.power_drift_warn_db} / fail {th.power_drift_fail_db}",
         "МНК-наклон по времени, нормированный на сутки; худшая линия."),
        ("Перекос линий, дБ", f"warn {th.lane_imbalance_warn_db} / fail {th.lane_imbalance_fail_db}",
         "Разброс средней мощности между 4 линиями QSFP28."),
        ("Дрейф Bias, мА/сут",
         f"warn {th.bias_drift_warn_ma_per_day} / fail {th.bias_drift_fail_ma_per_day}",
         "Лавинообразный рост тока смещения — признак уставшего лазера."),
        ("Чувствительность Bias", f"> {th.bias_per_degc_warn} мА/°C",
         "Рост тока, не объяснимый прогревом."),
        ("Запас до границы даташита", f"< {th.dom_margin_warn_db} дБ",
         "Малый запас до НИЖНЕЙ границы — предупреждение (нет запаса по бюджету)."),
        ("Минимум для оценки дрейфа", f"{th.min_hours_for_drift:g} ч",
         "Короче — наклон не экстраполируется на сутки, критерий не применяется."),
        ("Минимум снимков для перекоса", f"{th.min_samples_for_dom_stats}",
         "Меньше — перекос между линиями не оценивается."),
        ("Минимум наработки для ЗИП", f"{th.min_hours_for_pass:g} ч",
         "Короче — положительный вердикт не выносится: отсутствие ошибок за "
         "малый срок ничего не доказывает."),
        ("Жёсткие счётчики", "весь захват",
         "FEC Uncorrected, флапы, input/CRC/bit errors суммируются по всему "
         "захвату и не могут быть обнулены ни одним окном или фильтром."),
        ("Скорости и BER", "стабильные интервалы",
         "Считаются только там, где линк Up, нет флапов, сбросов и дефектов."),
        ("Классификация интервалов", "по логу",
         "STABLE / FLAP / RESET / LINK_DOWN / GAP — определяется состоянием линка "
         "и счётчиков, а не предположениями о действиях оператора."),
        ("Краевое окно на старте", f"{context['head_guard']:g} мин",
         "Исключается ТОЛЬКО из статистики DOM: снимок с вынутым модулем даёт "
         "ложный выход RX за границы даташита. На счётчики ошибок не влияет."),
        ("Краевое окно на финише", f"{context['tail_guard']:g} мин",
         "То же для конца захвата. На счётчики ошибок не влияет."),
        ("Обработка сбросов счётчиков", "автоматическая",
         "Отрицательная дельта = сброс; приростом считается значение после сброса, "
         "интервал помечается категорией RESET и остаётся в учёте."),
        ("Обработка пропусков опроса", "автоматическая",
         "Интервал длиннее 3 медиан помечается категорией GAP."),
    ]
    return pd.DataFrame(rows, columns=["Параметр", "Значение", "Комментарий"])


def generate_charts(
    df_sockets: pd.DataFrame,
    df_dom_raw: pd.DataFrame,
    df_intervals: pd.DataFrame,
    profile: OpticProfile,
    th: Thresholds,
    charts_dir: Path,
) -> None:
    """Строит все графики отчёта и складывает их в charts_dir."""
    charts_dir.mkdir(parents=True, exist_ok=True)
    setup_matplotlib()

    chart_ber_overview(df_sockets, th, charts_dir / "overview_ber.png")

    for _, r in df_sockets.iterrows():
        sid = r["socket_id"]
        slug = re.sub(r"[^A-Za-z0-9]+", "_", sid).strip("_")
        title = (f"Круг {r['test']} · {r['switch']} · порт {r['port']}"
                 + (f" · S/N {r['serial']}" if r.get("serial_is_real") else ""))

        dom = (df_dom_raw[(df_dom_raw["test"] == r["test"])
                          & (df_dom_raw["switch"] == r["switch"])
                          & (df_dom_raw["port"] == r["port"])]
               if not df_dom_raw.empty else pd.DataFrame())
        if not dom.empty:
            chart_thermal(dom, title, profile, charts_dir / f"{slug}_thermal.png")
            chart_optical_power(dom, title, profile, charts_dir / f"{slug}_power.png")

        iv = (df_intervals[df_intervals["socket_id"] == sid]
              if not df_intervals.empty else pd.DataFrame())
        if not iv.empty:
            chart_fec_timeline(iv, title, charts_dir / f"{slug}_fec.png")

    LOG.info("Графики построены: %s", charts_dir)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Разбор аргументов командной строки."""
    p = argparse.ArgumentParser(
        prog="sfp_soak_analyzer.py",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Анализ back-to-back soak-теста трансиверов QSFP28 на Juniper QFX.",
        epilog=textwrap.dedent("""\
            Примеры:
              %(prog)s
              %(prog)s --log-dir ./logs --out-dir ./report --optic SR4
              %(prog)s --map modules_map.csv --tail-guard-min 20
              %(prog)s --write-map-template   # создать шаблон карты ротации
        """),
    )
    p.add_argument("--log-dir", type=Path, default=Path("."),
                   help="Каталог с файлами логов (по умолчанию: текущий).")
    p.add_argument("--pattern", action="append", default=None,
                   help="Glob-маска файлов логов; можно указать несколько раз. "
                        "По умолчанию: '*sw-*.log', '*sw_*.log', '*.log'.")
    p.add_argument("--out-dir", type=Path, default=Path("./report"),
                   help="Каталог для результатов (по умолчанию: ./report).")
    p.add_argument("--map", dest="map_path", type=Path, default=None,
                   help="CSV-карта ротации: test,switch,port,serial[,partner_switch,partner_port].")
    p.add_argument("--write-map-template", action="store_true",
                   help="Создать шаблон карты ротации по обнаруженным портам и выйти.")
    p.add_argument("--optic", choices=sorted(OPTIC_PROFILES), default="SR4",
                   help="Профиль оптики для границ даташита (по умолчанию: SR4).")
    p.add_argument("--head-guard-min", type=float, default=5.0,
                   help="Краевое окно в начале захвата, мин (по умолчанию: 5).")
    p.add_argument("--tail-guard-min", type=float, default=15.0,
                   help="Краевое окно в конце захвата, мин (по умолчанию: 15).")
    p.add_argument("--ber-pass-max", type=float, default=None,
                   help="Порог pre-FEC BER для вердикта «в ЗИП».")
    p.add_argument("--ber-warn-max", type=float, default=None,
                   help="Порог pre-FEC BER, выше которого выносится брак.")
    p.add_argument("--no-charts", action="store_true",
                   help="Не строить графики (быстрый прогон, только Excel).")
    p.add_argument("--no-pdf", action="store_true", help="Не формировать PDF.")
    p.add_argument("--no-excel", action="store_true", help="Не формировать Excel.")
    p.add_argument("-v", "--verbose", action="store_true", help="Подробный лог.")
    return p.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Точка входа: разбор -> метрики -> вердикты -> отчёты."""
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)-8s %(message)s", stream=sys.stdout)

    log_dir: Path = args.log_dir
    if not log_dir.exists():
        LOG.error("Каталог логов не найден: %s", log_dir)
        return 2

    patterns = args.pattern or ["*sw-*.log", "*sw_*.log", "*.log"]

    # ------------------------------------------------------------ 1. разбор ----
    parser = SoakLogParser()
    captures = parser.discover(log_dir, patterns)
    if not captures:
        LOG.error("В каталоге %s не найдено ни одного подходящего файла логов "
                  "(маски: %s).", log_dir, ", ".join(patterns))
        return 2

    LOG.info("Найдено файлов логов: %d", len(captures))
    for c in captures:
        LOG.info("  %-22s -> круг %d, сторона %s", c.path.name, c.test, c.side)

    df_ports, df_dom_raw = parser.parse_all(captures)
    if df_ports.empty:
        LOG.error("Не удалось извлечь ни одной метрики портов — анализ невозможен.")
        return 3

    # ------------------------------------------------------ 2. карта модулей ---
    discovered = sorted({(int(r.test), str(r.switch), str(r.port))
                         for r in df_ports.itertuples()})
    if args.write_map_template:
        tpl = (args.map_path or (log_dir / "modules_map.csv"))
        ModuleMap.write_template(tpl, discovered)
        LOG.info("Шаблон создан. Заполните колонку 'serial' и запустите анализ "
                 "с ключом --map %s", tpl)
        return 0

    module_map = ModuleMap.load(args.map_path)
    if not module_map.active:
        LOG.warning("=" * 78)
        LOG.warning("Серийные номера недоступны: в выводе `show interfaces extensive` и")
        LOG.warning("`show interfaces diagnostics optics` их нет, карта --map не задана.")
        LOG.warning("Анализ будет выполнен по посадочным местам (круг/свитч/порт);")
        LOG.warning("кросс-раундовая локализация «модуль vs патч-корд» — недоступна.")
        LOG.warning("Создайте карту: %s --write-map-template", Path(sys.argv[0]).name)
        LOG.warning("=" * 78)

    # ---------------------------------------------------------- 3. метрики -----
    profile = OPTIC_PROFILES[args.optic]
    th = Thresholds()
    if args.ber_pass_max is not None:
        th.ber_pass_max = args.ber_pass_max
    if args.ber_warn_max is not None:
        th.ber_warn_max = args.ber_warn_max

    # Краевые окна размечаются и для счётчиков, и для DOM: снимок с извлечённым
    # модулем иначе даст ложный выход RX за границы даташита.
    df_ports = mark_edge_windows(df_ports, args.head_guard_min, args.tail_guard_min)
    if not df_dom_raw.empty:
        df_dom_raw = mark_edge_windows(df_dom_raw, args.head_guard_min,
                                       args.tail_guard_min)
    df_intervals = compute_intervals(df_ports)
    df_err = summarize_error_metrics(df_ports, df_intervals)
    df_events = build_error_events(df_intervals)
    df_dom_summary = summarize_dom(df_dom_raw, profile, th)

    # ------------------------------------------------------ 4. вердикты --------
    df_sockets = evaluate_all_sockets(df_err, df_dom_summary, profile, th)
    df_sockets["serial"] = [module_map.serial_for(r.test, r.switch, r.port)
                            for r in df_sockets.itertuples()]
    df_sockets["serial_is_real"] = [module_map.is_real(r.test, r.switch, r.port)
                                    for r in df_sockets.itertuples()]

    df_links = build_link_view(df_sockets, module_map)
    df_modules = apply_rotation_logic(df_sockets, module_map)

    # ------------------------------------------------------ 5. контекст --------
    ts_min = df_ports["ts"].min()
    ts_max = df_ports["ts"].max()
    poll = df_intervals["duration_s"].median() if not df_intervals.empty else None
    tz_labels = sorted({str(t) for t in df_ports["tz"].dropna().unique()})
    total_soak = float(pd.to_numeric(df_sockets["soak_seconds"],
                                     errors="coerce").fillna(0).sum())

    context: Dict[str, Any] = {
        "profile": profile,
        "thresholds": th,
        "map_active": module_map.active,
        "head_guard": args.head_guard_min,
        "tail_guard": args.tail_guard_min,
        "n_tests": int(df_sockets["test"].nunique()) if not df_sockets.empty else 0,
        "files_str": ", ".join(c.path.name for c in captures),
        "period_str": (f"{ts_min:%d.%m.%Y %H:%M} — {ts_max:%d.%m.%Y %H:%M}"
                       if pd.notna(ts_min) and pd.notna(ts_max) else "—"),
        "poll_interval_str": (f"{poll/60:.0f} мин" if poll and math.isfinite(poll) else "—"),
        "total_soak_str": human_duration(total_soak),
        "counter_resets": int(df_intervals["counter_reset"].sum())
                          if not df_intervals.empty else 0,
        "tz_conflict": " / ".join(tz_labels) if len(tz_labels) > 1 else None,
        "report_title": "Приёмочный soak-тест QSFP28 · Juniper QFX5110-48S-4C",
    }
    context["error_events"] = df_events
    context["methodology_df"] = build_methodology_table(profile, th, context)

    # ------------------------------------------------------ 6. отчёты ----------
    out_dir: Path = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = out_dir / "charts"

    if not args.no_charts:
        generate_charts(df_sockets, df_dom_raw, df_intervals, profile, th, charts_dir)

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    if not args.no_excel:
        write_excel(out_dir / f"qsfp28_soak_report_{stamp}.xlsx",
                    df_modules, df_sockets, df_dom_summary, df_links,
                    df_intervals, df_dom_raw, parser.issues, context)
    if not args.no_pdf:
        build_pdf(out_dir / f"qsfp28_soak_report_{stamp}.pdf",
                  df_modules, df_sockets, df_dom_summary, df_links,
                  df_intervals, df_dom_raw, parser.issues, context, charts_dir)

    # ------------------------------------------------------ 7. итог в консоль --
    print()
    print("=" * 86)
    print("ИТОГ АНАЛИЗА")
    print("=" * 86)
    counts = df_modules["final_verdict"].value_counts().to_dict()
    for verdict in (VERDICT_PASS, VERDICT_WARN, VERDICT_FAIL, VERDICT_UNKNOWN):
        if counts.get(verdict):
            print(f"  {verdict:<24} {counts[verdict]:>3}")
    print("-" * 86)
    header = f"{'Место / S/N':<26} {'Вердикт':<22} {'FEC Unc.':>9} " \
             f"{'Флапы':>6} {'pre-FEC BER':>12}"
    print(header)
    print("-" * 86)
    for _, r in df_modules.iterrows():
        print(f"  {str(r['serial'])[:24]:<24} {r['final_verdict']:<22} "
              f"{fmt_count(r['fec_uncorrected_total']):>9} "
              f"{fmt_count(r['link_flaps_total']):>6} "
              f"{fmt_ber(r['ber_worst']):>12}")
    print("=" * 86)
    print(f"Отчёты: {out_dir.resolve()}")
    if parser.issues:
        print(f"Замечаний при разборе: {len(parser.issues)} "
              f"(см. лист «Диагностика парсинга»)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
