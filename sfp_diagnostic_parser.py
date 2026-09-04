#!/usr/bin/env python3
"""
SFP Module Diagnostic Parser and Report Generator
Parses switch logs and generates comprehensive reports with graphs, CSV, TXT, and PDF outputs.
"""

import re
import os
import csv
from datetime import datetime, timezone, timedelta
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple
import statistics

# Data visualization and report generation
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_pdf import PdfPages
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT

@dataclass
class InterfaceMetrics:
    """Data class for interface metrics at a single point in time"""
    timestamp: datetime
    interface: str
    link_status: str = ""
    carrier_transitions: int = 0
    input_errors: int = 0
    drops: int = 0
    bit_errors: int = 0
    errored_blocks: int = 0
    fec_corrected: int = 0
    fec_uncorrected: int = 0
    crc_align_errors: int = 0
    temperature: Optional[float] = None  # Celsius
    temperature_alarm_high: bool = False
    temperature_alarm_low: bool = False
    temperature_warning_high: bool = False
    temperature_warning_low: bool = False
    temp_threshold_high_alarm: Optional[float] = None
    temp_threshold_low_alarm: Optional[float] = None
    temp_threshold_high_warn: Optional[float] = None
    temp_threshold_low_warn: Optional[float] = None
    bias_current: Optional[float] = None  # mA
    bias_current_alarm_high: bool = False
    bias_current_alarm_low: bool = False
    bias_current_warning_high: bool = False
    bias_current_warning_low: bool = False
    bias_threshold_high_alarm: Optional[float] = None
    bias_threshold_low_alarm: Optional[float] = None
    bias_threshold_high_warn: Optional[float] = None
    bias_threshold_low_warn: Optional[float] = None
    output_power_mw: Optional[float] = None
    output_power_dbm: Optional[float] = None
    output_power_alarm_high: bool = False
    output_power_alarm_low: bool = False
    output_power_warning_high: bool = False
    output_power_warning_low: bool = False
    output_threshold_high_alarm: Optional[float] = None
    output_threshold_low_alarm: Optional[float] = None
    output_threshold_high_warn: Optional[float] = None
    output_threshold_low_warn: Optional[float] = None
    rx_power_mw: Optional[float] = None
    rx_power_dbm: Optional[float] = None
    rx_power_alarm_high: bool = False
    rx_power_alarm_low: bool = False
    rx_power_warning_high: bool = False
    rx_power_warning_low: bool = False
    rx_threshold_high_alarm: Optional[float] = None
    rx_threshold_low_alarm: Optional[float] = None
    rx_threshold_high_warn: Optional[float] = None
    rx_threshold_low_warn: Optional[float] = None


@dataclass
class PortInfo:
    """Information about SFP module on a port"""
    port: str
    part_number: str = ""
    type_: str = "SR4"
    position: str = ""  # A48, B48, etc.
    issue: str = ""


# SFP module configuration based on user input
SFP_CONFIG_TEST1 = {
    "et-0/0/48": PortInfo("et-0/0/48", "QGA9882482", "SR4", "A48", ""),
    "et-0/0/49": PortInfo("et-0/0/49", "QGA9882272", "SR4", "B48", ""),
    "et-0/0/50": PortInfo("et-0/0/50", "QGA9882273", "SR4", "A49", ""),
    "et-0/0/51": PortInfo("et-0/0/51", "QGA9882479", "SR4", "B49", ""),
}

# Ports with issues in test 1
SFP_CONFIG_TEST1["et-0/0/52"] = PortInfo("et-0/0/52", "4020090014", "SR4", "A50", "Неправильно отображается описание (LR вместо SR)")
SFP_CONFIG_TEST1["et-0/0/53"] = PortInfo("et-0/0/53", "4020090024", "SR4", "B50", "Не отображается модель (UNKNOWN)")
SFP_CONFIG_TEST1["et-0/0/54"] = PortInfo("et-0/0/54", "4020090047", "SR4", "A51", "Неправильно отображается описание (LR вместо SR)")
SFP_CONFIG_TEST1["et-0/0/55"] = PortInfo("et-0/0/55", "6721070018", "SR4", "B51", "")

# Test 2 configuration - mapping ports based on user input
# Note: The actual port mapping may need adjustment based on actual hardware configuration
SFP_CONFIG_TEST2 = {
    "et-0/0/48": PortInfo("et-0/0/48", "4021040058", "SR4", "A48", ""),
    "et-0/0/49": PortInfo("et-0/0/49", "4021040060", "SR4", "B49", ""),
    "et-0/0/50": PortInfo("et-0/0/50", "4021040059", "SR4", "A50", ""),
    "et-0/0/51": PortInfo("et-0/0/51", "4021040063", "SR4", "A49", ""),
}

# Additional modules for test 2 with issues
SFP_CONFIG_TEST2["et-0/0/52"] = PortInfo("et-0/0/52", "4020090006", "SR4", "B48", "Не отображается модель (UNKNOWN)")
SFP_CONFIG_TEST2["et-0/0/53"] = PortInfo("et-0/0/53", "4021040057", "SR4", "B50", "")


class LogParser:
    """Parser for switch diagnostic logs"""
    
    def __init__(self):
        self.data: Dict[str, List[InterfaceMetrics]] = defaultdict(list)
        self.timestamps: List[datetime] = []
        
    def parse_timestamp(self, line: str) -> Optional[datetime]:
        """Extract timestamp from log line"""
        match = re.search(r'=====\s*(.+?)\s*=====', line)
        if match:
            ts_str = match.group(1).strip()
            # Manual parsing for format: "Tue Sep  1 16:26:19 +05 2026" or "Tue Sep  1 16:26:42 UTC 2026"
            
            # Try format with timezone offset: "Tue Sep  1 16:26:19 +05 2026"
            ts_match = re.match(r'(\w+)\s+(\w+)\s+(\d+)\s+(\d+):(\d+):(\d+)\s+([+-])(\d+)\s+(\d+)', ts_str)
            if ts_match:
                dow, month, day, hour, minute, second, tz_sign, tz_hour, year = ts_match.groups()
                months = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
                tz_offset = int(tz_hour)
                if tz_sign == '-':
                    tz_offset = -tz_offset
                
                tz = timezone(timedelta(hours=tz_offset))
                return datetime(int(year), months.get(month, 1), int(day), int(hour), int(minute), int(second), tzinfo=tz)
            
            # Try format with UTC: "Tue Sep  1 16:26:42 UTC 2026"
            ts_match_utc = re.match(r'(\w+)\s+(\w+)\s+(\d+)\s+(\d+):(\d+):(\d+)\s+(\w+)\s+(\d+)', ts_str)
            if ts_match_utc:
                dow, month, day, hour, minute, second, tz_name, year = ts_match_utc.groups()
                months = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
                tz = timezone.utc if tz_name == 'UTC' else timezone(timedelta(0))
                return datetime(int(year), months.get(month, 1), int(day), int(hour), int(minute), int(second), tzinfo=tz)
            
            # Try standard formats as fallback
            try:
                return datetime.strptime(ts_str, "%a %b %d %H:%M:%S %z %Y")
            except ValueError:
                pass
        return None
    
    def parse_interface_block(self, lines: List[str], current_ts: datetime) -> Dict[str, InterfaceMetrics]:
        """Parse a block of interface data"""
        metrics = {}
        current_interface = None
        lane_index = 0
        
        for line in lines:
            line = line.strip()
            
            # Detect interface
            iface_match = re.search(r'Physical interface:\s*(\S+),', line)
            if iface_match:
                current_interface = iface_match.group(1)
                lane_index = 0
                if current_interface not in metrics:
                    metrics[current_interface] = InterfaceMetrics(
                        timestamp=current_ts,
                        interface=current_interface
                    )
                # Parse link status
                if "Physical link is Up" in line:
                    metrics[current_interface].link_status = "Up"
                elif "Physical link is Down" in line:
                    metrics[current_interface].link_status = "Down"
                continue
            
            if current_interface is None:
                continue
                
            m = metrics[current_interface]
            
            # Parse input errors
            if "Carrier transitions:" in line:
                err_match = re.search(r'Carrier transitions:\s*(\d+),\s*Errors:\s*(\d+),\s*Drops:\s*(\d+)', line)
                if err_match:
                    m.carrier_transitions = int(err_match.group(1))
                    m.input_errors = int(err_match.group(2))
                    m.drops = int(err_match.group(3))
            
            # Parse bit errors
            if "Bit errors" in line and "threshold" not in line:
                match = re.search(r'Bit errors\s+(\d+)', line)
                if match:
                    m.bit_errors = int(match.group(1))
            
            # Parse errored blocks
            if "Errored blocks" in line:
                match = re.search(r'Errored blocks\s+(\d+)', line)
                if match:
                    m.errored_blocks = int(match.group(1))
            
            # Parse FEC statistics
            if "FEC Corrected Errors" in line and "Rate" not in line:
                match = re.search(r'FEC Corrected Errors\s+(\d+)', line)
                if match:
                    m.fec_corrected = int(match.group(1))
            
            if "FEC Uncorrected Errors" in line and "Rate" not in line:
                match = re.search(r'FEC Uncorrected Errors\s+(\d+)', line)
                if match:
                    m.fec_uncorrected = int(match.group(1))
            
            # Parse CRC errors
            if "CRC/Align errors" in line:
                match = re.search(r'CRC/Align errors\s+(\d+)\s+(\d+)', line)
                if match:
                    m.crc_align_errors = int(match.group(1))
            
            # Parse temperature
            if "Module temperature" in line and ":" in line and "threshold" not in line and "alarm" not in line and "warning" not in line:
                match = re.search(r'Module temperature\s*:\s*(\d+)\s*degrees C', line)
                if match:
                    m.temperature = float(match.group(1))
            
            # Parse temperature alarms/warnings
            if "Module temperature high alarm" in line and "threshold" not in line:
                m.temperature_alarm_high = "On" in line
            if "Module temperature low alarm" in line and "threshold" not in line:
                m.temperature_alarm_low = "On" in line
            if "Module temperature high warning" in line and "threshold" not in line:
                m.temperature_warning_high = "On" in line
            if "Module temperature low warning" in line and "threshold" not in line:
                m.temperature_warning_low = "On" in line
            
            # Parse temperature thresholds
            if "Module temperature high alarm threshold" in line:
                match = re.search(r'high alarm threshold\s*:\s*(\d+)\s*degrees C', line)
                if match:
                    m.temp_threshold_high_alarm = float(match.group(1))
            if "Module temperature low alarm threshold" in line:
                match = re.search(r'low alarm threshold\s*:\s*(-?\d+)\s*degrees C', line)
                if match:
                    m.temp_threshold_low_alarm = float(match.group(1))
            if "Module temperature high warning threshold" in line:
                match = re.search(r'high warning threshold\s*:\s*(\d+)\s*degrees C', line)
                if match:
                    m.temp_threshold_high_warn = float(match.group(1))
            if "Module temperature low warning threshold" in line:
                match = re.search(r'low warning threshold\s*:\s*(-?\d+)\s*degrees C', line)
                if match:
                    m.temp_threshold_low_warn = float(match.group(1))
            
            # Parse laser bias current (multiple lanes per port for QSFP)
            if "Laser bias current" in line and ":" in line and "threshold" not in line and "alarm" not in line and "warning" not in line:
                match = re.search(r'Laser bias current\s*:\s*([\d.]+)\s*mA', line)
                if match:
                    # For multi-lane modules, we might get multiple readings
                    # Store the latest or average
                    if m.bias_current is None:
                        m.bias_current = float(match.group(1))
                    else:
                        # Average multiple lanes
                        m.bias_current = (m.bias_current + float(match.group(1))) / 2
            
            # Parse bias current alarms/warnings
            if "Laser bias current high alarm" in line and "threshold" not in line:
                m.bias_current_alarm_high = "On" in line
            if "Laser bias current low alarm" in line and "threshold" not in line:
                m.bias_current_alarm_low = "On" in line
            if "Laser bias current high warning" in line and "threshold" not in line:
                m.bias_current_warning_high = "On" in line
            if "Laser bias current low warning" in line and "threshold" not in line:
                m.bias_current_warning_low = "On" in line
            
            # Parse bias current thresholds
            if "Laser bias current high alarm threshold" in line:
                match = re.search(r'high alarm threshold\s*:\s*([\d.]+)\s*mA', line)
                if match:
                    m.bias_threshold_high_alarm = float(match.group(1))
            if "Laser bias current low alarm threshold" in line:
                match = re.search(r'low alarm threshold\s*:\s*([\d.]+)\s*mA', line)
                if match:
                    m.bias_threshold_low_alarm = float(match.group(1))
            if "Laser bias current high warning threshold" in line:
                match = re.search(r'high warning threshold\s*:\s*([\d.]+)\s*mA', line)
                if match:
                    m.bias_threshold_high_warn = float(match.group(1))
            if "Laser bias current low warning threshold" in line:
                match = re.search(r'low warning threshold\s*:\s*([\d.]+)\s*mA', line)
                if match:
                    m.bias_threshold_low_warn = float(match.group(1))
            
            # Parse laser output power
            if "Laser output power" in line and ":" in line and "threshold" not in line and "alarm" not in line and "warning" not in line:
                mw_match = re.search(r'Laser output power\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if mw_match:
                    if m.output_power_mw is None:
                        m.output_power_mw = float(mw_match.group(1))
                        m.output_power_dbm = float(mw_match.group(2))
                    else:
                        # Average multiple lanes
                        m.output_power_mw = (m.output_power_mw + float(mw_match.group(1))) / 2
                        m.output_power_dbm = (m.output_power_dbm + float(mw_match.group(2))) / 2
            
            # Parse output power alarms
            if "Laser output power high alarm" in line and "threshold" not in line:
                m.output_power_alarm_high = "On" in line
            if "Laser output power low alarm" in line and "threshold" not in line:
                m.output_power_alarm_low = "On" in line
            if "Laser output power high warning" in line and "threshold" not in line:
                m.output_power_warning_high = "On" in line
            if "Laser output power low warning" in line and "threshold" not in line:
                m.output_power_warning_low = "On" in line
            
            # Parse output power thresholds
            if "Laser output power high alarm threshold" in line:
                match = re.search(r'high alarm threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.output_threshold_high_alarm = (float(match.group(1)), float(match.group(2)))
            if "Laser output power low alarm threshold" in line:
                match = re.search(r'low alarm threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.output_threshold_low_alarm = (float(match.group(1)), float(match.group(2)))
            if "Laser output power high warning threshold" in line:
                match = re.search(r'high warning threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.output_threshold_high_warn = (float(match.group(1)), float(match.group(2)))
            if "Laser output power low warning threshold" in line:
                match = re.search(r'low warning threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.output_threshold_low_warn = (float(match.group(1)), float(match.group(2)))
            
            # Parse laser receiver power
            if "Laser receiver power" in line and ":" in line and "threshold" not in line and "alarm" not in line and "warning" not in line:
                mw_match = re.search(r'Laser receiver power\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if mw_match:
                    if m.rx_power_mw is None:
                        m.rx_power_mw = float(mw_match.group(1))
                        m.rx_power_dbm = float(mw_match.group(2))
                    else:
                        # Average multiple lanes
                        m.rx_power_mw = (m.rx_power_mw + float(mw_match.group(1))) / 2
                        m.rx_power_dbm = (m.rx_power_dbm + float(mw_match.group(2))) / 2
            
            # Parse RX power alarms
            if "Laser receiver power high alarm" in line and "threshold" not in line:
                m.rx_power_alarm_high = "On" in line
            if "Laser receiver power low alarm" in line and "threshold" not in line:
                m.rx_power_alarm_low = "On" in line
            if "Laser receiver power high warning" in line and "threshold" not in line:
                m.rx_power_warning_high = "On" in line
            if "Laser receiver power low warning" in line and "threshold" not in line:
                m.rx_power_warning_low = "On" in line
            
            # Parse RX power thresholds
            if "Laser rx power high alarm threshold" in line:
                match = re.search(r'high alarm threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.rx_threshold_high_alarm = (float(match.group(1)), float(match.group(2)))
            if "Laser rx power low alarm threshold" in line:
                match = re.search(r'low alarm threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.rx_threshold_low_alarm = (float(match.group(1)), float(match.group(2)))
            if "Laser rx power high warning threshold" in line:
                match = re.search(r'high warning threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.rx_threshold_high_warn = (float(match.group(1)), float(match.group(2)))
            if "Laser rx power low warning threshold" in line:
                match = re.search(r'low warning threshold\s*:\s*([\d.]+)\s*mW\s*/\s*([-\d.]+)\s*dBm', line)
                if match:
                    m.rx_threshold_low_warn = (float(match.group(1)), float(match.group(2)))
        
        return metrics
    
    def parse_file(self, filepath: str) -> Dict[str, List[InterfaceMetrics]]:
        """Parse a complete log file"""
        data: Dict[str, List[InterfaceMetrics]] = defaultdict(list)
        
        # Read file as binary and decode to handle any encoding issues
        with open(filepath, 'rb') as f:
            raw_content = f.read()
        
        # Try to decode, replacing problematic characters
        try:
            content = raw_content.decode('utf-8', errors='replace')
        except:
            content = raw_content.decode('latin-1', errors='replace')
        
        # Split by timestamp markers
        blocks = re.split(r'(={5,}\s*.+?\s*={5,})', content)
        
        current_ts = None
        i = 0
        while i < len(blocks):
            block = blocks[i]
            
            # Check if this is a timestamp block
            if '=====' in block:
                current_ts = self.parse_timestamp(block)
                i += 1
                continue
            
            if current_ts and block.strip():
                # Parse interface data from this block
                lines = block.split('\n')
                metrics_dict = self.parse_interface_block(lines, current_ts)
                
                for iface, metrics in metrics_dict.items():
                    data[iface].append(metrics)
            
            i += 1
        
        return data


class DiagnosticAnalyzer:
    """Analyzer for SFP diagnostic data"""
    
    # SFP+ SR4 standard thresholds (typical values)
    TEMP_HIGH_ALARM = 75.0
    TEMP_LOW_ALARM = -5.0
    TEMP_HIGH_WARNING = 70.0
    TEMP_LOW_WARNING = 0.0
    
    def __init__(self, data: Dict[str, List[InterfaceMetrics]], port_info: Dict[str, PortInfo]):
        self.data = data
        self.port_info = port_info
        self.issues: List[Dict] = []
        
    def analyze(self) -> Dict:
        """Perform comprehensive analysis"""
        results = {
            'summary': {},
            'per_port': {},
            'issues': [],
            'recommendations': []
        }
        
        for port, metrics_list in self.data.items():
            if not metrics_list:
                continue
            
            port_result = self._analyze_port(port, metrics_list)
            results['per_port'][port] = port_result
            
            # Collect issues
            if port_result['issues']:
                for issue in port_result['issues']:
                    issue['port'] = port
                    results['issues'].append(issue)
        
        # Generate summary
        results['summary'] = self._generate_summary(results['per_port'])
        
        return results
    
    def _analyze_port(self, port: str, metrics_list: List[InterfaceMetrics]) -> Dict:
        """Analyze a single port"""
        result = {
            'measurements': len(metrics_list),
            'uptime_percentage': 0,
            'errors_detected': False,
            'issues': [],
            'statistics': {},
            'alarms': {
                'temperature': {'high': 0, 'low': 0},
                'bias_current': {'high': 0, 'low': 0},
                'output_power': {'high': 0, 'low': 0},
                'rx_power': {'high': 0, 'low': 0}
            }
        }
        
        if not metrics_list:
            return result
        
        # Extract time series data
        timestamps = [m.timestamp for m in metrics_list]
        temperatures = [m.temperature for m in metrics_list if m.temperature is not None]
        bias_currents = [m.bias_current for m in metrics_list if m.bias_current is not None]
        output_powers_dbm = [m.output_power_dbm for m in metrics_list if m.output_power_dbm is not None]
        rx_powers_dbm = [m.rx_power_dbm for m in metrics_list if m.rx_power_dbm is not None]
        link_up_count = sum(1 for m in metrics_list if m.link_status == "Up")
        
        # Calculate uptime
        result['uptime_percentage'] = (link_up_count / len(metrics_list)) * 100 if metrics_list else 0
        
        # Calculate statistics
        stats = {}
        if temperatures:
            stats['temperature'] = {
                'min': min(temperatures),
                'max': max(temperatures),
                'avg': statistics.mean(temperatures),
                'std': statistics.stdev(temperatures) if len(temperatures) > 1 else 0
            }
        if bias_currents:
            stats['bias_current_mA'] = {
                'min': min(bias_currents),
                'max': max(bias_currents),
                'avg': statistics.mean(bias_currents),
                'std': statistics.stdev(bias_currents) if len(bias_currents) > 1 else 0
            }
        if output_powers_dbm:
            stats['output_power_dbm'] = {
                'min': min(output_powers_dbm),
                'max': max(output_powers_dbm),
                'avg': statistics.mean(output_powers_dbm),
                'std': statistics.stdev(output_powers_dbm) if len(output_powers_dbm) > 1 else 0
            }
        if rx_powers_dbm:
            stats['rx_power_dbm'] = {
                'min': min(rx_powers_dbm),
                'max': max(rx_powers_dbm),
                'avg': statistics.mean(rx_powers_dbm),
                'std': statistics.stdev(rx_powers_dbm) if len(rx_powers_dbm) > 1 else 0
            }
        
        result['statistics'] = stats
        
        # Count alarms
        for m in metrics_list:
            if m.temperature_alarm_high:
                result['alarms']['temperature']['high'] += 1
            if m.temperature_alarm_low:
                result['alarms']['temperature']['low'] += 1
            if m.bias_current_alarm_high:
                result['alarms']['bias_current']['high'] += 1
            if m.bias_current_alarm_low:
                result['alarms']['bias_current']['low'] += 1
            if m.output_power_alarm_high:
                result['alarms']['output_power']['high'] += 1
            if m.output_power_alarm_low:
                result['alarms']['output_power']['low'] += 1
            if m.rx_power_alarm_high:
                result['alarms']['rx_power']['high'] += 1
            if m.rx_power_alarm_low:
                result['alarms']['rx_power']['low'] += 1
        
        # Detect issues
        if result['uptime_percentage'] < 100:
            result['issues'].append({
                'type': 'link_instability',
                'severity': 'high',
                'description': f'Link instability detected. Uptime: {result["uptime_percentage"]:.2f}%'
            })
        
        # Check for error counters increasing
        total_bit_errors = max((m.bit_errors for m in metrics_list), default=0)
        total_fec_uncorrected = max((m.fec_uncorrected for m in metrics_list), default=0)
        total_crc_errors = max((m.crc_align_errors for m in metrics_list), default=0)
        
        if total_bit_errors > 0:
            result['issues'].append({
                'type': 'bit_errors',
                'severity': 'medium',
                'description': f'Bit errors detected: {total_bit_errors}'
            })
            result['errors_detected'] = True
        
        if total_fec_uncorrected > 0:
            result['issues'].append({
                'type': 'fec_uncorrected',
                'severity': 'high',
                'description': f'Uncorrectable FEC errors: {total_fec_uncorrected}'
            })
            result['errors_detected'] = True
        
        if total_crc_errors > 0:
            result['issues'].append({
                'type': 'crc_errors',
                'severity': 'medium',
                'description': f'CRC/Alignment errors: {total_crc_errors}'
            })
            result['errors_detected'] = True
        
        # Check temperature issues
        if stats.get('temperature'):
            if stats['temperature']['max'] > self.TEMP_HIGH_WARNING:
                result['issues'].append({
                    'type': 'temperature_warning',
                    'severity': 'medium',
                    'description': f'Temperature exceeded warning threshold: {stats["temperature"]["max"]:.1f}°C'
                })
            if stats['temperature']['max'] > self.TEMP_HIGH_ALARM:
                result['issues'].append({
                    'type': 'temperature_alarm',
                    'severity': 'high',
                    'description': f'Temperature exceeded alarm threshold: {stats["temperature"]["max"]:.1f}°C'
                })
        
        # Check RX power levels (critical for optical modules)
        if stats.get('rx_power_dbm'):
            avg_rx = stats['rx_power_dbm']['avg']
            # SR4 typical receive sensitivity: -7.4 dBm to 2.4 dBm
            if avg_rx < -7.4:
                result['issues'].append({
                    'type': 'low_rx_power',
                    'severity': 'high',
                    'description': f'Received optical power too low: {avg_rx:.2f} dBm (threshold: -7.4 dBm)'
                })
            elif avg_rx > 2.4:
                result['issues'].append({
                    'type': 'high_rx_power',
                    'severity': 'medium',
                    'description': f'Received optical power too high: {avg_rx:.2f} dBm (threshold: 2.4 dBm)'
                })
        
        # Add SFP module info if available
        if port in self.port_info:
            sfp = self.port_info[port]
            if sfp.issue:
                result['issues'].append({
                    'type': 'module_info',
                    'severity': 'info',
                    'description': f'SFP Issue: {sfp.issue}'
                })
        
        return result
    
    def _generate_summary(self, per_port_results: Dict) -> Dict:
        """Generate overall summary"""
        total_ports = len(per_port_results)
        ports_with_issues = sum(1 for p in per_port_results.values() if p['issues'])
        total_alarms = sum(
            sum(a['high'] + a['low'] for a in p['alarms'].values())
            for p in per_port_results.values()
        )
        
        return {
            'total_ports': total_ports,
            'ports_with_issues': ports_with_issues,
            'healthy_ports': total_ports - ports_with_issues,
            'total_alarms': total_alarms,
            'overall_health': 'Good' if ports_with_issues == 0 else 'Issues Detected'
        }


class ReportGenerator:
    """Generate various report formats"""
    
    def __init__(self, analysis_results: Dict, raw_data: Dict[str, List[InterfaceMetrics]], 
                 port_info: Dict[str, PortInfo], test_name: str, output_dir: str):
        self.results = analysis_results
        self.raw_data = raw_data
        self.port_info = port_info
        self.test_name = test_name
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
    
    def generate_all(self):
        """Generate all report formats"""
        print(f"Generating reports for {self.test_name}...")
        self.generate_csv()
        self.generate_txt_report()
        self.generate_plots()
        self.generate_pdf_report()
        print(f"Reports saved to: {self.output_dir}")
    
    def generate_csv(self):
        """Generate CSV files with raw data"""
        for port, metrics_list in self.raw_data.items():
            if not metrics_list:
                continue
            
            filename = os.path.join(self.output_dir, f"{self.test_name}_{port.replace('/', '_')}_data.csv")
            with open(filename, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow([
                    'Timestamp', 'Link Status', 'Temperature (C)', 'Bias Current (mA)',
                    'Output Power (mW)', 'Output Power (dBm)', 'RX Power (mW)', 'RX Power (dBm)',
                    'Bit Errors', 'FEC Corrected', 'FEC Uncorrected', 'CRC Errors'
                ])
                for m in metrics_list:
                    writer.writerow([
                        m.timestamp.isoformat() if m.timestamp else '',
                        m.link_status,
                        m.temperature if m.temperature else '',
                        m.bias_current if m.bias_current else '',
                        m.output_power_mw if m.output_power_mw else '',
                        m.output_power_dbm if m.output_power_dbm else '',
                        m.rx_power_mw if m.rx_power_mw else '',
                        m.rx_power_dbm if m.rx_power_dbm else '',
                        m.bit_errors,
                        m.fec_corrected,
                        m.fec_uncorrected,
                        m.crc_align_errors
                    ])
        print(f"  CSV files generated")
    
    def generate_txt_report(self):
        """Generate text report"""
        filename = os.path.join(self.output_dir, f"{self.test_name}_report.txt")
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write(f"SFP MODULE DIAGNOSTIC REPORT - {self.test_name}\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            # Summary
            summary = self.results['summary']
            f.write("SUMMARY\n")
            f.write("-" * 40 + "\n")
            f.write(f"Total Ports Analyzed: {summary['total_ports']}\n")
            f.write(f"Healthy Ports: {summary['healthy_ports']}\n")
            f.write(f"Ports with Issues: {summary['ports_with_issues']}\n")
            f.write(f"Total Alarms: {summary['total_alarms']}\n")
            f.write(f"Overall Health Status: {summary['overall_health']}\n\n")
            
            # Per-port details
            f.write("PER-PORT ANALYSIS\n")
            f.write("-" * 40 + "\n\n")
            
            for port, port_result in sorted(self.results['per_port'].items()):
                f.write(f"Port: {port}\n")
                
                # SFP Info
                if port in self.port_info:
                    sfp = self.port_info[port]
                    f.write(f"  Part Number: {sfp.part_number}\n")
                    f.write(f"  Type: {sfp.type_}\n")
                    f.write(f"  Position: {sfp.position}\n")
                    if sfp.issue:
                        f.write(f"  Known Issue: {sfp.issue}\n")
                
                f.write(f"  Measurements: {port_result['measurements']}\n")
                f.write(f"  Uptime: {port_result['uptime_percentage']:.2f}%\n")
                
                # Statistics
                if port_result['statistics']:
                    f.write("  Statistics:\n")
                    for metric, stats in port_result['statistics'].items():
                        f.write(f"    {metric}: min={stats['min']:.3f}, max={stats['max']:.3f}, ")
                        f.write(f"avg={stats['avg']:.3f}, std={stats['std']:.3f}\n")
                
                # Alarms
                total_alarms = sum(a['high'] + a['low'] for a in port_result['alarms'].values())
                if total_alarms > 0:
                    f.write("  Alarms:\n")
                    for alarm_type, counts in port_result['alarms'].items():
                        if counts['high'] or counts['low']:
                            f.write(f"    {alarm_type}: high={counts['high']}, low={counts['low']}\n")
                
                # Issues
                if port_result['issues']:
                    f.write("  Issues Detected:\n")
                    for issue in port_result['issues']:
                        f.write(f"    [{issue['severity'].upper()}] {issue['description']}\n")
                
                f.write("\n")
            
            # All issues summary
            if self.results['issues']:
                f.write("\nALL ISSUES SUMMARY\n")
                f.write("-" * 40 + "\n")
                for issue in self.results['issues']:
                    f.write(f"[{issue.get('port', 'N/A')}] [{issue['severity'].upper()}] {issue['description']}\n")
            
            # Recommendations
            f.write("\n\nRECOMMENDATIONS\n")
            f.write("-" * 40 + "\n")
            if not self.results['issues']:
                f.write("No critical issues detected. Continue regular monitoring.\n")
            else:
                high_severity = [i for i in self.results['issues'] if i['severity'] == 'high']
                if high_severity:
                    f.write("HIGH PRIORITY:\n")
                    for issue in high_severity:
                        f.write(f"  - {issue['port']}: {issue['description']}\n")
                    f.write("  Recommended action: Immediate investigation required\n\n")
                
                medium_severity = [i for i in self.results['issues'] if i['severity'] == 'medium']
                if medium_severity:
                    f.write("MEDIUM PRIORITY:\n")
                    for issue in medium_severity:
                        f.write(f"  - {issue['port']}: {issue['description']}\n")
                    f.write("  Recommended action: Schedule maintenance window\n")
        
        print(f"  Text report generated")
    
    def generate_plots(self):
        """Generate matplotlib plots"""
        plt.style.use('default')
        
        # Create figure directory
        fig_dir = os.path.join(self.output_dir, 'figures')
        os.makedirs(fig_dir, exist_ok=True)
        
        for port, metrics_list in self.raw_data.items():
            if not metrics_list or len(metrics_list) < 2:
                continue
            
            # Extract data
            timestamps = [m.timestamp for m in metrics_list if m.timestamp]
            if not timestamps:
                continue
            
            temperatures = [m.temperature for m in metrics_list]
            bias_currents = [m.bias_current for m in metrics_list]
            output_powers = [m.output_power_dbm for m in metrics_list]
            rx_powers = [m.rx_power_dbm for m in metrics_list]
            
            # Create multi-panel plot
            fig, axes = plt.subplots(4, 1, figsize=(14, 12))
            fig.suptitle(f'SFP Diagnostics - Port {port} - {self.test_name}', fontsize=14, fontweight='bold')
            
            # Temperature plot
            ax = axes[0]
            if any(t is not None for t in temperatures):
                valid_idx = [i for i, t in enumerate(temperatures) if t is not None]
                ax.plot([timestamps[i] for i in valid_idx], [temperatures[i] for i in valid_idx], 
                       'r-', label='Temperature', linewidth=1)
                ax.axhline(y=70, color='orange', linestyle='--', alpha=0.7, label='Warning (70°C)')
                ax.axhline(y=75, color='red', linestyle='--', alpha=0.7, label='Alarm (75°C)')
                ax.set_ylabel('Temperature (°C)')
                ax.legend(loc='upper right')
                ax.grid(True, alpha=0.3)
            
            # Bias Current plot
            ax = axes[1]
            if any(b is not None for b in bias_currents):
                valid_idx = [i for i, b in enumerate(bias_currents) if b is not None]
                ax.plot([timestamps[i] for i in valid_idx], [bias_currents[i] for i in valid_idx], 
                       'b-', label='Bias Current', linewidth=1)
                ax.set_ylabel('Bias Current (mA)')
                ax.legend(loc='upper right')
                ax.grid(True, alpha=0.3)
            
            # Output Power plot
            ax = axes[2]
            if any(p is not None for p in output_powers):
                valid_idx = [i for i, p in enumerate(output_powers) if p is not None]
                ax.plot([timestamps[i] for i in valid_idx], [output_powers[i] for i in valid_idx], 
                       'g-', label='Output Power', linewidth=1)
                ax.axhline(y=-7.6, color='orange', linestyle='--', alpha=0.7, label='Warning Low')
                ax.axhline(y=2.4, color='orange', linestyle='--', alpha=0.7, label='Warning High')
                ax.set_ylabel('Output Power (dBm)')
                ax.legend(loc='upper right')
                ax.grid(True, alpha=0.3)
            
            # RX Power plot
            ax = axes[3]
            if any(p is not None for p in rx_powers):
                valid_idx = [i for i, p in enumerate(rx_powers) if p is not None]
                ax.plot([timestamps[i] for i in valid_idx], [rx_powers[i] for i in valid_idx], 
                       'm-', label='RX Power', linewidth=1)
                ax.axhline(y=-10.3, color='orange', linestyle='--', alpha=0.7, label='Warning Low')
                ax.axhline(y=-7.4, color='red', linestyle='--', alpha=0.7, label='Min Sensitivity')
                ax.axhline(y=2.4, color='orange', linestyle='--', alpha=0.7, label='Max Input')
                ax.set_ylabel('RX Power (dBm)')
                ax.set_xlabel('Time')
                ax.legend(loc='upper right')
                ax.grid(True, alpha=0.3)
            
            # Format x-axis dates
            for ax in axes:
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M\n%m/%d'))
                ax.xaxis.set_major_locator(mdates.AutoDateLocator())
                plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            plt.tight_layout()
            plt.savefig(os.path.join(fig_dir, f'{port.replace("/", "_")}_diagnostics.png'), 
                       dpi=150, bbox_inches='tight')
            plt.close()
        
        # Create comparison plot for all ports
        fig, axes = plt.subplots(2, 2, figsize=(16, 10))
        fig.suptitle(f'All Ports Comparison - {self.test_name}', fontsize=14, fontweight='bold')
        
        colors_port = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
        
        for idx, (port, metrics_list) in enumerate(sorted(self.raw_data.items())):
            if not metrics_list:
                continue
            color = colors_port[idx % len(colors_port)]
            port_short = port.replace('et-0/0/', '')
            
            # Temperature
            ax = axes[0, 0]
            temps = [(m.timestamp, m.temperature) for m in metrics_list if m.temperature]
            if temps:
                ax.plot([t[0] for t in temps], [t[1] for t in temps], color=color, 
                       label=f'Port {port_short}', linewidth=1, alpha=0.8)
            ax.set_ylabel('Temperature (°C)')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            
            # RX Power
            ax = axes[0, 1]
            rx = [(m.timestamp, m.rx_power_dbm) for m in metrics_list if m.rx_power_dbm is not None]
            if rx:
                ax.plot([t[0] for t in rx], [t[1] for t in rx], color=color, 
                       label=f'Port {port_short}', linewidth=1, alpha=0.8)
            ax.set_ylabel('RX Power (dBm)')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            
            # Bias Current
            ax = axes[1, 0]
            bias = [(m.timestamp, m.bias_current) for m in metrics_list if m.bias_current is not None]
            if bias:
                ax.plot([t[0] for t in bias], [t[1] for t in bias], color=color, 
                       label=f'Port {port_short}', linewidth=1, alpha=0.8)
            ax.set_ylabel('Bias Current (mA)')
            ax.set_xlabel('Time')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            
            # Output Power
            ax = axes[1, 1]
            out = [(m.timestamp, m.output_power_dbm) for m in metrics_list if m.output_power_dbm is not None]
            if out:
                ax.plot([t[0] for t in out], [t[1] for t in out], color=color, 
                       label=f'Port {port_short}', linewidth=1, alpha=0.8)
            ax.set_ylabel('Output Power (dBm)')
            ax.set_xlabel('Time')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
        
        for ax in axes.flatten():
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M\n%m/%d'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        plt.tight_layout()
        plt.savefig(os.path.join(fig_dir, 'all_ports_comparison.png'), dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"  Plots generated in {fig_dir}")
    
    def generate_pdf_report(self):
        """Generate comprehensive PDF report"""
        filename = os.path.join(self.output_dir, f"{self.test_name}_full_report.pdf")
        
        doc = SimpleDocTemplate(filename, pagesize=landscape(A4),
                               rightMargin=2*cm, leftMargin=2*cm,
                               topMargin=2*cm, bottomMargin=2*cm)
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            textColor=colors.darkblue,
            spaceAfter=12,
            alignment=TA_CENTER
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            textColor=colors.darkblue,
            spaceAfter=10,
            spaceBefore=10
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            alignment=TA_LEFT
        )
        
        story = []
        
        # Title
        story.append(Paragraph(f"SFP Module Diagnostic Report", title_style))
        story.append(Paragraph(f"Test: {self.test_name}", heading_style))
        story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
        story.append(Spacer(1, 0.3*inch))
        
        # Summary table
        summary = self.results['summary']
        story.append(Paragraph("Executive Summary", heading_style))
        
        summary_data = [
            ['Metric', 'Value'],
            ['Total Ports', str(summary['total_ports'])],
            ['Healthy Ports', str(summary['healthy_ports'])],
            ['Ports with Issues', str(summary['ports_with_issues'])],
            ['Total Alarms', str(summary['total_alarms'])],
            ['Overall Health', summary['overall_health']]
        ]
        
        summary_table = Table(summary_data, colWidths=[3*inch, 2*inch])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 0.3*inch))
        
        # SFP Module Information
        story.append(Paragraph("SFP Module Configuration", heading_style))
        sfp_data = [['Port', 'Part Number', 'Type', 'Position', 'Known Issues']]
        for port in sorted(self.port_info.keys()):
            sfp = self.port_info[port]
            sfp_data.append([
                port,
                sfp.part_number,
                sfp.type_,
                sfp.position,
                sfp.issue if sfp.issue else 'None'
            ])
        
        sfp_table = Table(sfp_data, colWidths=[1.2*inch, 1.5*inch, 0.8*inch, 0.8*inch, 2*inch])
        sfp_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(sfp_table)
        story.append(PageBreak())
        
        # Per-port detailed analysis
        story.append(Paragraph("Per-Port Detailed Analysis", heading_style))
        
        for port, port_result in sorted(self.results['per_port'].items()):
            story.append(Paragraph(f"Port: {port}", heading_style))
            
            # SFP info for this port
            if port in self.port_info:
                sfp = self.port_info[port]
                port_info_text = f"<b>Part Number:</b> {sfp.part_number} | <b>Type:</b> {sfp.type_} | <b>Position:</b> {sfp.position}"
                if sfp.issue:
                    port_info_text += f"<br/><b>Issue:</b> <font color='red'>{sfp.issue}</font>"
                story.append(Paragraph(port_info_text, normal_style))
            
            # Key metrics
            metrics_table_data = [
                ['Metric', 'Value'],
                ['Measurements', str(port_result['measurements'])],
                ['Uptime', f"{port_result['uptime_percentage']:.2f}%"],
                ['Errors Detected', 'Yes' if port_result['errors_detected'] else 'No']
            ]
            
            # Add statistics
            for metric_name, stats in port_result['statistics'].items():
                metrics_table_data.append([
                    f"{metric_name} (avg)",
                    f"{stats['avg']:.3f} ± {stats['std']:.3f}"
                ])
            
            metrics_table = Table(metrics_table_data, colWidths=[2*inch, 2*inch])
            metrics_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ]))
            story.append(metrics_table)
            
            # Issues for this port
            if port_result['issues']:
                story.append(Spacer(1, 0.1*inch))
                story.append(Paragraph("<b>Issues Detected:</b>", normal_style))
                for issue in port_result['issues']:
                    severity_color = {'high': 'red', 'medium': 'orange', 'low': 'blue', 'info': 'green'}.get(issue['severity'], 'black')
                    story.append(Paragraph(
                        f"<font color='{severity_color}'>[{issue['severity'].upper()}]</font> {issue['description']}",
                        normal_style
                    ))
            
            story.append(Spacer(1, 0.2*inch))
        
        # Add plots
        story.append(PageBreak())
        story.append(Paragraph("Diagnostic Graphs", heading_style))
        
        fig_dir = os.path.join(self.output_dir, 'figures')
        if os.path.exists(fig_dir):
            for port in sorted(self.raw_data.keys()):
                plot_file = os.path.join(fig_dir, f'{port.replace("/", "_")}_diagnostics.png')
                if os.path.exists(plot_file):
                    story.append(Paragraph(f"Diagnostics for Port {port}", heading_style))
                    # Scale image to fit page
                    img = Image(plot_file, width=9*inch, height=6*inch)
                    story.append(img)
                    story.append(Spacer(1, 0.2*inch))
            
            # Add comparison plot
            comparison_plot = os.path.join(fig_dir, 'all_ports_comparison.png')
            if os.path.exists(comparison_plot):
                story.append(PageBreak())
                story.append(Paragraph("All Ports Comparison", heading_style))
                img = Image(comparison_plot, width=9*inch, height=5*inch)
                story.append(img)
        
        # Recommendations
        story.append(PageBreak())
        story.append(Paragraph("Recommendations", heading_style))
        
        if not self.results['issues']:
            story.append(Paragraph("No critical issues detected. Continue regular monitoring.", normal_style))
        else:
            high_issues = [i for i in self.results['issues'] if i['severity'] == 'high']
            medium_issues = [i for i in self.results['issues'] if i['severity'] == 'medium']
            
            if high_issues:
                story.append(Paragraph("<b>HIGH PRIORITY ACTIONS:</b>", normal_style))
                for issue in high_issues:
                    story.append(Paragraph(f"• {issue.get('port', 'N/A')}: {issue['description']}", normal_style))
                story.append(Spacer(1, 0.1*inch))
            
            if medium_issues:
                story.append(Paragraph("<b>MEDIUM PRIORITY ACTIONS:</b>", normal_style))
                for issue in medium_issues:
                    story.append(Paragraph(f"• {issue.get('port', 'N/A')}: {issue['description']}", normal_style))
        
        doc.build(story)
        print(f"  PDF report generated")


def main():
    """Main function to run the complete analysis"""
    base_dir = "/workspace"
    
    # File mappings
    test_files = {
        "Test1_swA": os.path.join(base_dir, "sw-a1.log"),
        "Test1_swB": os.path.join(base_dir, "sw-b1.log"),
        "Test2_swA": os.path.join(base_dir, "sw-A2.log"),
        "Test2_swB": os.path.join(base_dir, "sw-B2.log"),
    }
    
    # SFP configurations
    sfp_configs = {
        "Test1": SFP_CONFIG_TEST1,
        "Test2": SFP_CONFIG_TEST2,
    }
    
    parser = LogParser()
    
    for test_name, filepath in test_files.items():
        print(f"\n{'='*60}")
        print(f"Processing: {test_name}")
        print(f"File: {filepath}")
        print('='*60)
        
        if not os.path.exists(filepath):
            print(f"Warning: File not found: {filepath}")
            continue
        
        # Determine which test this is (Test1 or Test2)
        test_key = "Test1" if "a1" in filepath.lower() or "b1" in filepath.lower() else "Test2"
        port_info = sfp_configs[test_key]
        
        # Parse log file
        print("Parsing log file...")
        data = parser.parse_file(filepath)
        
        if not data:
            print("Warning: No data parsed from file")
            continue
        
        print(f"Parsed {len(data)} ports")
        for port, metrics in data.items():
            print(f"  {port}: {len(metrics)} measurements")
        
        # Analyze data
        print("Analyzing data...")
        analyzer = DiagnosticAnalyzer(data, port_info)
        results = analyzer.analyze()
        
        # Generate reports
        output_dir = os.path.join(base_dir, f"reports_{test_name}")
        generator = ReportGenerator(results, data, port_info, test_name, output_dir)
        generator.generate_all()
    
    print("\n" + "="*60)
    print("Processing complete!")
    print("="*60)


if __name__ == "__main__":
    main()
