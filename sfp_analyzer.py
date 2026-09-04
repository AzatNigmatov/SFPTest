#!/usr/bin/env python3
"""
SFP Module Diagnostic Log Analyzer for Juniper QFX switches
Analyzes soak test logs to identify trends, issues, and failing modules
"""

import re
import os
from datetime import datetime
from collections import defaultdict
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np

# Configuration
LOG_FILES = {
    'sw-A': '/workspace/sw-a1.log',
    'sw-B': '/workspace/sw-b1.log'
}

INTERFACES = ['et-0/0/48', 'et-0/0/49', 'et-0/0/50', 'et-0/0/51']

def read_log_file(filepath):
    """Read log file handling potential binary content"""
    import subprocess
    # Use strings to extract readable text from potentially binary log files
    result = subprocess.run(['strings', filepath], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout
    # Fallback: try reading as text
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception as e:
        print(f"Error reading file: {e}")
        return ""

def parse_timestamp(line):
    """Extract timestamp from log section header"""
    match = re.search(r'=====\s*(.+?)\s*=====', line)
    if match:
        ts_str = match.group(1).strip()
        
        # Remove UTC keyword if present
        ts_str = ts_str.replace(' UTC', '')
        
        # Normalize whitespace (handle double spaces in date format)
        ts_str = re.sub(r'\s+', ' ', ts_str)
        
        # Fix timezone format (+05 -> +0500)
        ts_str = re.sub(r'([+-]\d{2})\s', r'\g<1>00 ', ts_str)
        
        try:
            return datetime.strptime(ts_str, '%a %b %d %H:%M:%S %z %Y')
        except:
            pass
        
        # Try without timezone
        ts_no_tz = re.sub(r'\s*[+-]\d{2}\s*', ' ', match.group(1))
        ts_no_tz = ts_no_tz.replace(' UTC', '')
        ts_no_tz = re.sub(r'\s+', ' ', ts_no_tz).strip()
        try:
            return datetime.strptime(ts_no_tz, '%a %b %d %H:%M:%S %Y')
        except:
            pass
    return None

def extract_value(text, pattern, group=1):
    """Extract numeric value using regex pattern"""
    match = re.search(pattern, text)
    if match:
        try:
            return float(match.group(group))
        except:
            return None
    return None

def parse_interface_section(section, interface_name):
    """Parse a single interface section from the log"""
    data = {
        'interface': interface_name,
        'link_status': None,
        'carrier_transitions': None,
        'input_errors': None,
        'bit_errors': None,
        'errored_blocks': None,
        'fec_corrected': None,
        'fec_uncorrected': None,
        'crc_errors': None,
        'temperature': None,
        'temp_high_alarm': None,
        'temp_low_alarm': None,
        'temp_high_warn': None,
        'temp_low_warn': None,
        'bias_current': None,
        'bias_high_alarm': None,
        'bias_low_alarm': None,
        'tx_power_mw': None,
        'tx_power_dbm': None,
        'rx_power_mw': None,
        'rx_power_dbm': None,
        'rx_high_alarm': None,
        'rx_low_alarm': None,
        'rx_high_warn': None,
        'rx_low_warn': None,
    }
    
    # Link status
    if 'Physical link is Up' in section:
        data['link_status'] = 'Up'
    elif 'Physical link is Down' in section:
        data['link_status'] = 'Down'
    
    # Carrier transitions
    match = re.search(r'Carrier transitions:\s*(\d+)', section)
    if match:
        data['carrier_transitions'] = int(match.group(1))
    
    # Input errors
    match = re.search(r'Input errors:.*?Errors:\s*(\d+)', section, re.DOTALL)
    if match:
        data['input_errors'] = int(match.group(1))
    
    # Bit errors
    match = re.search(r'Bit errors\s+(\d+)', section)
    if match:
        data['bit_errors'] = int(match.group(1))
    
    # Errored blocks
    match = re.search(r'Errored blocks\s+(\d+)', section)
    if match:
        data['errored_blocks'] = int(match.group(1))
    
    # FEC errors
    match = re.search(r'FEC Corrected Errors\s+(?:Rate\s+)?(\d+)', section)
    if match:
        data['fec_corrected'] = int(match.group(1))
    
    match = re.search(r'FEC Uncorrected Errors\s+(?:Rate\s+)?(\d+)', section)
    if match:
        data['fec_uncorrected'] = int(match.group(1))
    
    # CRC errors
    match = re.search(r'CRC/Align errors\s+(\d+)', section)
    if match:
        data['crc_errors'] = int(match.group(1))
    
    # Temperature
    temp_match = re.search(r'Module temperature\s+:\s*(-?\d+(?:\.\d+)?)\s*degrees C', section)
    if temp_match:
        data['temperature'] = float(temp_match.group(1))
    
    # Temperature alarms/warnings
    data['temp_high_alarm'] = 'On' if re.search(r'Module temperature high alarm\s+:\s*On', section) else 'Off'
    data['temp_low_alarm'] = 'On' if re.search(r'Module temperature low alarm\s+:\s*On', section) else 'Off'
    data['temp_high_warn'] = 'On' if re.search(r'Module temperature high warning\s+:\s*On', section) else 'Off'
    data['temp_low_warn'] = 'On' if re.search(r'Module temperature low warning\s+:\s*On', section) else 'Off'
    
    # Laser bias current (take the last reading in section)
    bias_matches = re.findall(r'Laser bias current\s+:\s*(-?\d+(?:\.\d+)?)\s*mA', section)
    if bias_matches:
        data['bias_current'] = float(bias_matches[-1])
    
    # Bias current alarms
    data['bias_high_alarm'] = 'On' if re.search(r'Laser bias current high alarm\s+:\s*On', section) else 'Off'
    data['bias_low_alarm'] = 'On' if re.search(r'Laser bias current low alarm\s+:\s*On', section) else 'Off'
    
    # TX power (last reading)
    tx_mw_matches = re.findall(r'Laser output power\s+:\s*(-?\d+(?:\.\d+)?)\s*mW', section)
    tx_dbm_matches = re.findall(r'Laser output power\s+:\s*[^\n]*?/\s*(-?\d+(?:\.\d+)?)\s*dBm', section)
    if tx_mw_matches:
        data['tx_power_mw'] = float(tx_mw_matches[-1])
    if tx_dbm_matches:
        data['tx_power_dbm'] = float(tx_dbm_matches[-1])
    
    # RX power (last reading)
    rx_mw_matches = re.findall(r'Laser receiver power\s+:\s*(-?\d+(?:\.\d+)?)\s*mW', section)
    rx_dbm_matches = re.findall(r'Laser receiver power\s+:\s*[^\n]*?/\s*(-?\d+(?:\.\d+)?)\s*dBm', section)
    if rx_mw_matches:
        data['rx_power_mw'] = float(rx_mw_matches[-1])
    if rx_dbm_matches:
        data['rx_power_dbm'] = float(rx_dbm_matches[-1])
    
    # RX alarms/warnings
    data['rx_high_alarm'] = 'On' if re.search(r'Laser receiver power high alarm\s+:\s*On', section) else 'Off'
    data['rx_low_alarm'] = 'On' if re.search(r'Laser receiver power low alarm\s+:\s*On', section) else 'Off'
    data['rx_high_warn'] = 'On' if re.search(r'Laser receiver power high warning\s+:\s*On', section) else 'Off'
    data['rx_low_warn'] = 'On' if re.search(r'Laser receiver power low warning\s+:\s*On', section) else 'Off'
    
    return data

def parse_log_content(content, switch_name):
    """Parse entire log content and return structured data"""
    records = []
    
    # Split by timestamp markers - handle both with and without trailing newline
    sections = re.split(r'(=====.*?=====\n?)', content)
    
    current_timestamp = None
    i = 0
    while i < len(sections):
        section = sections[i]
        
        # Check if this is a timestamp section
        if '=====' in section:
            current_timestamp = parse_timestamp(section)
            i += 1
            continue
        
        if current_timestamp and section.strip():
            # Parse each interface in this section
            for iface in INTERFACES:
                # Find interface section - look for the interface header
                iface_start = section.find(f'Physical interface: {iface}')
                if iface_start == -1:
                    continue
                
                # Find the end of this interface section (next interface or end)
                next_iface = section.find('Physical interface:', iface_start + 1)
                if next_iface == -1:
                    iface_section = section[iface_start:]
                else:
                    iface_section = section[iface_start:next_iface]
                
                iface_data = parse_interface_section(iface_section, iface)
                iface_data['timestamp'] = current_timestamp
                iface_data['switch'] = switch_name
                records.append(iface_data)
        
        i += 1
    
    return records

def analyze_data(df):
    """Perform comprehensive analysis on the data"""
    analysis = {}
    
    for switch in df['switch'].unique():
        switch_df = df[df['switch'] == switch]
        analysis[switch] = {}
        
        for iface in INTERFACES:
            iface_df = switch_df[switch_df['interface'] == iface].copy()
            
            if iface_df.empty:
                continue
            
            iface_analysis = {
                'samples': len(iface_df),
                'issues': [],
                'trends': {},
                'statistics': {}
            }
            
            # Calculate statistics for key metrics
            metrics = ['rx_power_dbm', 'tx_power_dbm', 'bias_current', 'temperature']
            for metric in metrics:
                if metric in iface_df.columns and iface_df[metric].notna().any():
                    values = iface_df[metric].dropna()
                    iface_analysis['statistics'][metric] = {
                        'min': values.min(),
                        'max': values.max(),
                        'mean': values.mean(),
                        'std': values.std(),
                        'trend': 'stable'
                    }
                    
                    # Detect trend
                    if len(values) > 10:
                        first_quarter = values.iloc[:len(values)//4].mean()
                        last_quarter = values.iloc[-len(values)//4:].mean()
                        change = last_quarter - first_quarter
                        
                        if metric in ['bias_current']:
                            if change > 0.5:
                                iface_analysis['statistics'][metric]['trend'] = 'increasing_degradation'
                                iface_analysis['issues'].append(f"Bias current increasing: {change:.2f} mA")
                            elif change < -0.5:
                                iface_analysis['statistics'][metric]['trend'] = 'decreasing'
                        elif metric in ['rx_power_dbm', 'tx_power_dbm']:
                            if change < -1.0:
                                iface_analysis['statistics'][metric]['trend'] = 'decreasing_power'
                                iface_analysis['issues'].append(f"Optical power decreasing: {change:.2f} dBm")
                            elif change > 1.0:
                                iface_analysis['statistics'][metric]['trend'] = 'increasing_power'
            
            # Check for error conditions
            if iface_df['fec_uncorrected'].sum() > 0:
                iface_analysis['issues'].append(f"FEC uncorrected errors: {iface_df['fec_uncorrected'].sum()}")
            
            if iface_df['crc_errors'].sum() > 0:
                iface_analysis['issues'].append(f"CRC errors detected: {iface_df['crc_errors'].sum()}")
            
            if iface_df['bit_errors'].sum() > 0:
                iface_analysis['issues'].append(f"Bit errors detected: {iface_df['bit_errors'].sum()}")
            
            # Check carrier transitions (link flaps)
            carrier_max = iface_df['carrier_transitions'].max()
            carrier_min = iface_df['carrier_transitions'].min()
            if carrier_max - carrier_min > 5:
                iface_analysis['issues'].append(f"Link instability: {carrier_max - carrier_min} carrier transitions")
            
            # Check alarms
            if (iface_df['rx_low_alarm'] == 'On').any():
                iface_analysis['issues'].append("RX power LOW ALARM triggered")
            if (iface_df['rx_high_alarm'] == 'On').any():
                iface_analysis['issues'].append("RX power HIGH ALARM triggered")
            if (iface_df['bias_high_alarm'] == 'On').any():
                iface_analysis['issues'].append("Bias current HIGH ALARM triggered")
            if (iface_df['temp_high_alarm'] == 'On').any():
                iface_analysis['issues'].append("Temperature HIGH ALARM triggered")
            
            # Check for link down events
            if (iface_df['link_status'] == 'Down').any():
                down_count = (iface_df['link_status'] == 'Down').sum()
                iface_analysis['issues'].append(f"Link went DOWN {down_count} times")
            
            # RX power quality assessment
            rx_values = iface_df['rx_power_dbm'].dropna()
            if len(rx_values) > 0:
                # Typical SFP+ RX sensitivity: -14 to -1 dBm for good signal
                # Below -14 dBm: weak signal
                # Above -1 dBm: potentially too strong
                weak_signal = (rx_values < -12).sum()
                strong_signal = (rx_values > -2).sum()
                
                if weak_signal > len(rx_values) * 0.5:
                    iface_analysis['issues'].append(f"Weak RX signal: {weak_signal}/{len(rx_values)} samples below -12 dBm")
                    iface_analysis['trends']['weak_signal_ratio'] = weak_signal / len(rx_values)
            
            analysis[switch][iface] = iface_analysis
    
    return analysis

def create_report(df, analysis, output_dir='/workspace'):
    """Generate comprehensive report with charts"""
    
    # Create output directory if needed
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate PDF report with charts
    pdf_path = os.path.join(output_dir, 'sfp_analysis_report.pdf')
    
    with PdfPages(pdf_path) as pdf:
        # Page 1: Summary
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.axis('off')
        
        summary_text = "SFP MODULE DIAGNOSTIC ANALYSIS REPORT\n"
        summary_text += "=" * 50 + "\n\n"
        summary_text += f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        summary_text += f"Switches Analyzed: {', '.join(df['switch'].unique())}\n"
        summary_text += f"Interfaces: {', '.join(INTERFACES)}\n"
        summary_text += f"Total Samples: {len(df)}\n\n"
        
        # Count issues
        total_issues = 0
        critical_issues = []
        for switch in analysis:
            for iface in analysis[switch]:
                issues = analysis[switch][iface].get('issues', [])
                total_issues += len(issues)
                for issue in issues:
                    if 'ALARM' in issue or 'DOWN' in issue or 'uncorrected' in issue.lower():
                        critical_issues.append(f"{switch} {iface}: {issue}")
        
        summary_text += f"Total Issues Found: {total_issues}\n"
        if critical_issues:
            summary_text += f"\nCRITICAL ISSUES ({len(critical_issues)}):\n"
            for issue in critical_issues[:10]:
                summary_text += f"  • {issue}\n"
        
        ax.text(0.1, 0.95, summary_text, transform=ax.transAxes, fontsize=11,
                verticalalignment='top', fontfamily='monospace')
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()
        
        # Pages for each switch
        for switch in df['switch'].unique():
            switch_df = df[df['switch'] == switch]
            
            # RX Power over time
            fig, axes = plt.subplots(2, 2, figsize=(14, 10))
            fig.suptitle(f'{switch} - SFP Diagnostics', fontsize=14, fontweight='bold')
            
            # Plot 1: RX Power
            ax = axes[0, 0]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty and iface_df['rx_power_dbm'].notna().any():
                    ax.plot(iface_df['timestamp'], iface_df['rx_power_dbm'], 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.axhline(y=-12, color='orange', linestyle='--', label='Weak signal threshold')
            ax.axhline(y=-14.29, color='red', linestyle=':', label='Low alarm threshold')
            ax.set_ylabel('RX Power (dBm)')
            ax.set_title('Received Optical Power')
            ax.legend(loc='lower right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # Plot 2: TX Power
            ax = axes[0, 1]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty and iface_df['tx_power_dbm'].notna().any():
                    ax.plot(iface_df['timestamp'], iface_df['tx_power_dbm'], 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.set_ylabel('TX Power (dBm)')
            ax.set_title('Transmitted Optical Power')
            ax.legend(loc='lower right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # Plot 3: Laser Bias Current
            ax = axes[1, 0]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty and iface_df['bias_current'].notna().any():
                    ax.plot(iface_df['timestamp'], iface_df['bias_current'], 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.axhline(y=12, color='orange', linestyle='--', label='High warning ~12mA')
            ax.set_ylabel('Bias Current (mA)')
            ax.set_title('Laser Bias Current')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # Plot 4: Temperature
            ax = axes[1, 1]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty and iface_df['temperature'].notna().any():
                    ax.plot(iface_df['timestamp'], iface_df['temperature'], 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.axhline(y=70, color='orange', linestyle='--', label='Warning threshold')
            ax.axhline(y=75, color='red', linestyle=':', label='Alarm threshold')
            ax.set_ylabel('Temperature (°C)')
            ax.set_title('Module Temperature')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
            
            # Error metrics page
            fig, axes = plt.subplots(2, 2, figsize=(14, 10))
            fig.suptitle(f'{switch} - Error Metrics', fontsize=14, fontweight='bold')
            
            # FEC Corrected Errors
            ax = axes[0, 0]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty:
                    ax.plot(iface_df['timestamp'], iface_df['fec_corrected'].fillna(0), 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.set_ylabel('FEC Corrected Errors')
            ax.set_title('Forward Error Correction - Corrected')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # FEC Uncorrected Errors
            ax = axes[0, 1]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty:
                    ax.plot(iface_df['timestamp'], iface_df['fec_uncorrected'].fillna(0), 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.set_ylabel('FEC Uncorrected Errors')
            ax.set_title('Forward Error Correction - Uncorrected (Critical)')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # CRC Errors
            ax = axes[1, 0]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty:
                    ax.plot(iface_df['timestamp'], iface_df['crc_errors'].fillna(0), 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.set_ylabel('CRC Errors')
            ax.set_title('Cyclic Redundancy Check Errors')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # Carrier Transitions
            ax = axes[1, 1]
            for iface in INTERFACES:
                iface_df = switch_df[switch_df['interface'] == iface]
                if not iface_df.empty:
                    ax.plot(iface_df['timestamp'], iface_df['carrier_transitions'].fillna(0), 
                           label=iface, marker='.', markersize=3, alpha=0.7)
            ax.set_ylabel('Carrier Transitions')
            ax.set_title('Link State Changes (Flaps)')
            ax.legend(loc='upper right', fontsize=8)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight')
            plt.close()
    
    # Generate CSV with all data
    csv_path = os.path.join(output_dir, 'sfp_detailed_data.csv')
    df.to_csv(csv_path, index=False)
    
    # Generate text summary
    txt_path = os.path.join(output_dir, 'sfp_analysis_summary.txt')
    with open(txt_path, 'w') as f:
        f.write("=" * 70 + "\n")
        f.write("SFP MODULE DIAGNOSTIC ANALYSIS - DETAILED SUMMARY\n")
        f.write("=" * 70 + "\n\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Log files analyzed: {list(LOG_FILES.keys())}\n")
        f.write(f"Total data points: {len(df)}\n\n")
        
        for switch in sorted(analysis.keys()):
            f.write("\n" + "=" * 70 + "\n")
            f.write(f"SWITCH: {switch}\n")
            f.write("=" * 70 + "\n\n")
            
            for iface in sorted(analysis[switch].keys()):
                data = analysis[switch][iface]
                f.write(f"\n--- Interface: {iface} ---\n")
                f.write(f"Samples collected: {data['samples']}\n")
                
                if data['statistics']:
                    f.write("\nStatistics:\n")
                    for metric, stats in data['statistics'].items():
                        f.write(f"  {metric}:\n")
                        f.write(f"    Min: {stats['min']:.3f}, Max: {stats['max']:.3f}, ")
                        f.write(f"Mean: {stats['mean']:.3f}, Std: {stats['std']:.3f}\n")
                        f.write(f"    Trend: {stats['trend']}\n")
                
                if data['issues']:
                    f.write("\n⚠️  ISSUES DETECTED:\n")
                    for issue in data['issues']:
                        f.write(f"  • {issue}\n")
                else:
                    f.write("\n✓ No issues detected\n")
                
                f.write("\n")
        
        # Overall recommendations
        f.write("\n" + "=" * 70 + "\n")
        f.write("RECOMMENDATIONS\n")
        f.write("=" * 70 + "\n\n")
        
        recommendations = []
        for switch in analysis:
            for iface in analysis[switch]:
                issues = analysis[switch][iface].get('issues', [])
                for issue in issues:
                    if 'ALARM' in issue:
                        recommendations.append(f"URGENT: Replace SFP on {switch} {iface} - {issue}")
                    elif 'uncorrected' in issue.lower():
                        recommendations.append(f"HIGH: Investigate {switch} {iface} - {issue}")
                    elif 'Weak RX' in issue:
                        recommendations.append(f"MEDIUM: Check fiber/cleaning on {switch} {iface} - {issue}")
                    elif 'Bias current increasing' in issue:
                        recommendations.append(f"MEDIUM: Monitor {switch} {iface} - laser degradation suspected")
        
        if recommendations:
            for rec in recommendations:
                f.write(f"• {rec}\n")
        else:
            f.write("• All modules appear healthy. Continue monitoring.\n")
    
    return pdf_path, csv_path, txt_path

def main():
    print("SFP Module Diagnostic Log Analyzer")
    print("=" * 50)
    
    all_records = []
    
    for switch_name, log_path in LOG_FILES.items():
        print(f"\nProcessing {switch_name}: {log_path}")
        
        if not os.path.exists(log_path):
            print(f"  Warning: File not found, skipping...")
            continue
        
        content = read_log_file(log_path)
        records = parse_log_content(content, switch_name)
        print(f"  Parsed {len(records)} records")
        all_records.extend(records)
    
    if not all_records:
        print("No data parsed. Exiting.")
        return
    
    # Create DataFrame
    df = pd.DataFrame(all_records)
    print(f"\nTotal records: {len(df)}")
    print(f"Columns: {df.columns.tolist()}")
    
    # Convert timestamp to datetime - handle mixed timezone-aware and naive
    df['timestamp'] = pd.to_datetime(df['timestamp'], utc=True)
    # Convert to naive datetime for consistent handling
    df['timestamp'] = df['timestamp'].dt.tz_localize(None)
    
    # Run analysis
    print("\nRunning analysis...")
    analysis = analyze_data(df)
    
    # Generate reports
    print("\nGenerating reports...")
    pdf_path, csv_path, txt_path = create_report(df, analysis)
    
    print(f"\nReports generated:")
    print(f"  PDF Report: {pdf_path}")
    print(f"  CSV Data: {csv_path}")
    print(f"  Text Summary: {txt_path}")
    
    # Print quick summary
    print("\n" + "=" * 50)
    print("QUICK SUMMARY")
    print("=" * 50)
    
    for switch in sorted(analysis.keys()):
        print(f"\n{switch}:")
        for iface in sorted(analysis[switch].keys()):
            issues = analysis[switch][iface].get('issues', [])
            status = "⚠️ ISSUES" if issues else "✓ OK"
            print(f"  {iface}: {status}")
            if issues:
                for issue in issues[:2]:
                    print(f"    - {issue}")

if __name__ == '__main__':
    main()
