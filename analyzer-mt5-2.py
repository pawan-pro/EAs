import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta #length wrt cntrl f time
import os
import time
import pytz
import MetaTrader5 as mt5
from mplfinance.original_flavor import candlestick_ohlc
from mpl_toolkits.axes_grid1 import make_axes_locatable
import re
from pptx import Presentation
from pptx.util import Inches

# Initialize MT5 connection
if not mt5.initialize():
    print("MT5 initialization failed")
    mt5.shutdown()

def estimate_typical_spread(symbol):
    """
    Estimate a typical spread for the given symbol.
    """
    # Major currency pairs
    major_fx = {
        "EURUSD.sd": 0.00010, "USDJPY.sd": 0.00010, "GBPUSD.sd": 0.00015,
        "USDCHF.sd": 0.00015, "AUDUSD.sd": 0.00015, "USDCAD.sd": 0.00015
    }
    
    # Minor currency pairs
    minor_fx = {
        "EURGBP.sd": 0.00020, "EURJPY.sd": 0.00020, "GBPJPY.sd": 0.00025,
        "EURCHF.sd": 0.00025, "AUDJPY.sd": 0.00020, "NZDUSD.sd": 0.00020
    }
    
    # Check if the symbol is in major or minor currency pairs
    if symbol in major_fx:
        return major_fx[symbol]
    elif symbol in minor_fx:
        return minor_fx[symbol]
    else:
        # For other symbols, return a default spread
        return 0.00030  # Default to 3 pips

def get_current_spread(symbol):
    symbol_info = mt5.symbol_info(symbol)
    if symbol_info is None:
        print(f"Failed to get symbol info for {symbol}")
        return None
    
    spread = symbol_info.ask - symbol_info.bid
    
    if spread == 0:
        typical_spread = estimate_typical_spread(symbol)
        print(f"MT5 returned 0 spread for {symbol}. Using estimated typical spread: {typical_spread}")
        return typical_spread
    
    return spread

def fetch_data(symbol, start_time, end_time, timeframe):
    # Convert times to UTC timezone
    timezone = pytz.timezone("Etc/UTC")
    start_time_tz = start_time.astimezone(timezone)
    end_time_tz = end_time.astimezone(timezone)

    # Request OHLC data from MT5
    rates = mt5.copy_rates_range(symbol, timeframe, start_time_tz, end_time_tz)
    
    if rates is not None and len(rates) > 0:
        print(f"Fetched {len(rates)} bars for {symbol} with timeframe {timeframe}")
        # Convert to DataFrame
        df = pd.DataFrame(rates)
        df['time'] = pd.to_datetime(df['time'], unit='s', utc=True)
        df = df.rename(columns={'open': 'open', 'high': 'high', 'low': 'low', 'close': 'close', 'tick_volume': 'volume'})

        df = calculate_atr(df)

        return df

    print(f"No OHLC data available for {symbol} with timeframe {timeframe}")
    return None

def calculate_atr(df, period=14):
    df['tr1'] = df['high'] - df['low']
    df['tr2'] = (df['high'] - df['close'].shift()).abs()
    df['tr3'] = (df['low'] - df['close'].shift()).abs()
    df['tr'] = df[['tr1', 'tr2', 'tr3']].max(axis=1)
    df['atr'] = df['tr'].rolling(window=period).mean()
    return df

import re

def parse_excel_input(excel_input):
    # Replace '^I' with '\t' to simulate actual tab characters
    excel_input = excel_input.replace('^I', '\t')
    lines = excel_input.strip().split('\n')
    events = []
    date_time_inputs = {}
    event_data_list = {}
    event_symbols = {}
    
    current_event_name = ""
    current_event_time = ""
    
    for line in lines:
        # Clean up extra spaces and tabs
        line = re.sub(r'\s{2,}', '\t', line.strip())  # Replace multiple spaces with a tab

        if line.startswith('Event'):
            parts = line.split('\t')
            current_event_name = parts[1].strip() if len(parts) > 1 else ""
            events.append({'name': current_event_name})
        elif line.startswith('Actual:'):
            parts = line.split('\t')
            symbol = parts[-1].strip() if len(parts) > 1 else ""
            event_symbols[current_event_name] = {"mt5": symbol}
        elif line.startswith('Time (GMT):'):
            parts = line.split('\t')
            current_event_time = parts[1].strip() if len(parts) > 1 else ""
        elif re.match(r'\d{2}-\w{3}-\d{2}', line):
            # Split by tab, assuming the cleaned line has been properly formatted
            parts = line.split('\t') 
            
            if len(parts) < 4:
                print(f"Skipping line due to insufficient data: {line}")
                continue  # Skip this line if there are not enough parts
            
            date_str = parts[0].strip()
            date = datetime.strptime(date_str, '%d-%b-%y').strftime('%Y-%m-%d')
            actual = parts[1].strip().replace('%', '') if len(parts) > 1 else "NA"
            forecast = parts[2].strip().replace('%', '') if len(parts) > 2 else "NA"
            previous = parts[3].strip().replace('%', '') if len(parts) > 3 else "NA"
            time_str = current_event_time or ""  # Use the event time if present, else leave empty

            if date not in date_time_inputs:
                date_time_inputs[date] = {
                    "date": date,
                    "event": current_event_time + ":00",
                }
                event_data_list[date] = []
            
            event_data_list[date].append({
                "name": current_event_name,
                "actual": actual,
                "forecast": forecast,
                "previous": previous,
                "time": time_str
            })
    
    return date_time_inputs, event_data_list, events, event_symbols

# Initialize presentation
template_path = "/Users/pawan/Desktop/Quantwater Tech Investments/Research & Development/Research/Blog/Slides/template.pptx"  # Update this path
prs = Presentation(template_path)
generated_image_filenames = []

def plot_ohlc(df, symbol, event_time, event_data, atr_multiple, spread, atr_period=14,
              is_last_chart=False, event_summary=None, start_time=None, end_time=None, timeframe=None):
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14.5, 9),
                                   gridspec_kw={'height_ratios': [4, 1]}, sharex=True)
    
    df['time_num'] = mdates.date2num(df['time'])
    ohlc = df[['time_num', 'open', 'high', 'low', 'close']].values

    plt.style.use('ggplot')

    candlestick_ohlc(ax1, ohlc, width=0.0004, colorup='green', colordown='red')

    # Adjust x-axis tick labels rotation and alignment
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    plt.setp(ax1.get_xticklabels(), rotation=45, ha='right')

    # Plot event time marker
    closest_index = (df['time'] - event_time).abs().idxmin()
    closest_time = df.loc[closest_index, 'time']
    event_price = df.loc[closest_index, 'open']

    ax1.plot(closest_time, event_price, 'X', color='red', markersize=10, label='Event Time')

    # Calculate volatility unit
    atr_value = df.loc[closest_index, 'atr']
    volatility_unit = atr_value * atr_multiple + spread

    # Plot volatility unit lines
    for i in range(1, 6):
        upper_range = event_price + i * volatility_unit
        lower_range = event_price - i * volatility_unit
        ax1.axhline(upper_range, color='blue', linestyle='--', label=f'+{i} Volatility Unit' if i == 1 else "")
        ax1.axhline(lower_range, color='orange', linestyle='--', label=f'-{i} Volatility Unit' if i == 1 else "")

    # Plot max and min peaks
    post_event_df = df[df['time'] >= event_time]
    max_price = post_event_df['high'].max()
    min_price = post_event_df['low'].min()
    max_movement = max_price - event_price
    min_movement = event_price - min_price

    # Check if there are rows where the conditions are met for max_time and min_time
    if not post_event_df[post_event_df['high'] == max_price].empty:
        max_time = post_event_df.loc[post_event_df['high'] == max_price, 'time'].iloc[0]
        ax1.plot(max_time, max_price, 'gx', markersize=10, label='Max Peak')

    if not post_event_df[post_event_df['low'] == min_price].empty:
        min_time = post_event_df.loc[post_event_df['low'] == min_price, 'time'].iloc[0]
        ax1.plot(min_time, min_price, 'rx', markersize=10, label='Min Peak')

    # Calculate pre-ATR volatility
    pre_atr_df = df[df['time'] < event_time].iloc[-(atr_period*2):-atr_period]
    pre_atr_volatility = pre_atr_df['high'].max() - pre_atr_df['low'].min()

    # Plot volume
    volume_color = 'blue'
    volume_alpha = 0.3
    ax2.bar(df['time'], df['volume'], width=0.0004, color=volume_color, alpha=volume_alpha)
    ax2.set_ylabel('Volume')

    # Prepare info text
    info_text = (
        f'ATR: {atr_value:.4f}\n'
        f'Spread: {spread:.5f}\n'
        f'Pre-ATR Volatility: {pre_atr_volatility:.4f}\n'
        f'\n'
        f'Volatility Unit (ATR*{atr_multiple}+spread): {volatility_unit:.4f}\n'
        f'Max Movement: {max_movement/volatility_unit:.2f}x\n'
        f'Min Movement: {min_movement/volatility_unit:.2f}x\n'
        f'Timeframe: {timeframe_to_string(timeframe)}'
    )

    if is_last_chart and event_summary and event_summary['Avg Volatility Unit'] > 0:
        info_text += (
            f'\n\nEvent Summary:\n'
            f'Avg Volatility Unit: {event_summary["Avg Volatility Unit"]:.4f}\n'
            f'Avg Max Movement: {event_summary["Avg Max Movement"]:.2f}x\n'
            f'Avg Min Movement: {event_summary["Avg Min Movement"]:.2f}x\n'
            f'Avg Pre-ATR Volatility: {event_summary["Avg Pre-ATR Volatility"]:.4f}'
        )

    # Add info text to the plot
    ax1.text(1.01, 0.5, info_text, transform=ax1.transAxes, ha='left', va='center', fontsize=9)

    # Set x-axis limits
    if start_time and end_time:
        ax1.set_xlim(mdates.date2num(start_time), mdates.date2num(end_time))

    # Create a new axis for event labels
    divider = make_axes_locatable(ax2)
    ax_event = divider.append_axes("bottom", size="15%", pad=0.1, sharex=ax1)
    ax_event.axis('off')  # Hide axis frame and ticks

    # Adjust the bottom margin to make room for labels
    plt.subplots_adjust(bottom=0.45)  # Increased from 0.35 to 0.45

    # Plot event labels
    for i, event in enumerate(event_data):
        event_name = event['name']
        actual = event['actual']
        forecast = event['forecast']
        previous = event['previous']

        # Determine label color
        if forecast == "NA":
            event_color = 'black'
        else:
            try:
                if float(actual) > float(forecast):
                    event_color = 'green'
                elif float(actual) < float(forecast):
                    event_color = 'red'
                else:
                    event_color = 'blue'
            except ValueError:
                event_color = 'black'

        event_time_num = mdates.date2num(event_time)

        # Combine event information into one line
        event_text_line = f"{event_name} | Actual: {actual} | Forecast: {forecast} | Previous: {previous}"

        # Calculate y position to prevent overlapping with x-axis tick labels
        y_offset = -0.2 - i * 0.1  # Adjusted y_offset to position labels further down
        ax_event.text(event_time_num, y_offset, event_text_line, color=event_color,
                      ha='center', va='top', fontsize=9, rotation=0, transform=ax_event.transData, fontweight='bold')

    ax1.set_title(f'{symbol} Price Action - {event_time.strftime("%Y-%m-%d %H:%M")} GMT + 3', fontweight='bold')
    ax1.set_ylabel('Price')
    ax1.grid(True)

    # Adjust legend
    ax1.legend(loc='upper left', bbox_to_anchor=(1, 1))

    plt.tight_layout()
    plt.subplots_adjust(hspace=0)

    # Save the figure as an image
    image_filename = f"{symbol}_{event_time.strftime('%Y%m%d_%H%M%S')}.png"
    plt.savefig(image_filename)
    plt.close(fig)

    return image_filename, volatility_unit, max_movement, min_movement, pre_atr_volatility

def timeframe_to_string(timeframe):
    timeframe_dict = {
        mt5.TIMEFRAME_M1: "M1",
        mt5.TIMEFRAME_M2: "M2",
        mt5.TIMEFRAME_M3: "M3",
        mt5.TIMEFRAME_M4: "M4",
        mt5.TIMEFRAME_M5: "M5",
        mt5.TIMEFRAME_M10: "M10",
        mt5.TIMEFRAME_M15: "M15",
        mt5.TIMEFRAME_M30: "M30",
        mt5.TIMEFRAME_H1: "H1",
        mt5.TIMEFRAME_H4: "H4",
        mt5.TIMEFRAME_D1: "D1",
        mt5.TIMEFRAME_W1: "W1",
        mt5.TIMEFRAME_MN1: "MN1"
    }
    return timeframe_dict.get(timeframe, f"{timeframe}")

# List of timeframes to try
timeframes = [
    mt5.TIMEFRAME_M1, mt5.TIMEFRAME_M2, mt5.TIMEFRAME_M3, mt5.TIMEFRAME_M4, mt5.TIMEFRAME_M5,
    mt5.TIMEFRAME_M10, mt5.TIMEFRAME_M15, mt5.TIMEFRAME_M30, mt5.TIMEFRAME_H1,
    mt5.TIMEFRAME_H4, mt5.TIMEFRAME_D1, mt5.TIMEFRAME_W1
]

def process_event(event, date_time_inputs, event_data_list, event_symbols):
    event_summary = {
        'volatility_units': [],
        'max_movements': [],
        'min_movements': [],
        'pre_atr_volatilities': [],
        'Avg Volatility Unit': 0,
        'Avg Max Movement': 0,
        'Avg Min Movement': 0,
        'Avg Pre-ATR Volatility': 0
    }

    event_dates = [
        date for date in sorted(date_time_inputs.keys())
        if any(e['name'] == event['name'] for e in event_data_list[date])
    ]

    for date in event_dates:
        event_time = datetime.strptime(date_time_inputs[date]["event"], '%H:%M:%S').time()
        event_datetime = datetime.combine(date_time_inputs[date]["date"], event_time)
        event_data = event_data_list[date]

        for event_info in event_data:
            if event_info['name'] == event['name']:
                symbol = event_symbols[event['name']]["mt5"]
                df = fetch_data(symbol, event_datetime - timedelta(hours=2), event_datetime + timedelta(hours=2))

                if df is not None:
                    image_filename, volatility_unit, max_movement, min_movement, pre_atr_volatility = plot_ohlc(
                        df, symbol, event_datetime, event_data, atr_multiple=3, spread=get_current_spread(symbol),
                        is_last_chart=(event_info == event_data[-1]), event_summary=event_summary,
                        start_time=event_datetime - timedelta(hours=2), end_time=event_datetime + timedelta(hours=2)
                    )

                    event_summary['volatility_units'].append(volatility_unit)
                    event_summary['max_movements'].append(max_movement)
                    event_summary['min_movements'].append(min_movement)
                    event_summary['pre_atr_volatilities'].append(pre_atr_volatility)

                    print(f"Processed {symbol} at {event_datetime.strftime('%Y-%m-%d %H:%M')} GMT")
                    generated_image_filenames.append(image_filename)

    if event_summary['volatility_units']:
        event_summary['Avg Volatility Unit'] = sum(event_summary['volatility_units']) / len(event_summary['volatility_units'])
        event_summary['Avg Max Movement'] = sum(event_summary['max_movements']) / len(event_summary['max_movements'])
        event_summary['Avg Min Movement'] = sum(event_summary['min_movements']) / len(event_summary['min_movements'])
        event_summary['Avg Pre-ATR Volatility'] = sum(event_summary['pre_atr_volatilities']) / len(event_summary['pre_atr_volatilities'])
    else:
        print(f"No data available for event: {event['name']}")

    return event_summary

def main():
    atr_multiple = 1.5
    atr_period = 14

    print("Please paste your Excel-formatted event data (press Enter twice when finished):")
    excel_input = ""
    while True:
        line = input()
        if line.strip() == "":
            break
        excel_input += line + "\n"

    date_time_inputs, event_data_list, events, event_symbols = parse_excel_input(excel_input)

    for event in events:
        event_summary = process_event(event, date_time_inputs, event_data_list, event_symbols)
        print(f"Event Summary for {event['name']}:")
        print(event_summary)

    output_pptx_path = "/Users/pawan/Desktop/Quantwater Tech Investments/Research & Development/Research/Blog/Slides/updated_presentation.pptx"
    prs.save(output_pptx_path)

    # Clean up temporary image files
    for image_filename in generated_image_filenames:
        os.remove(image_filename)

    # Shutdown MT5
    mt5.shutdown()

if __name__ == "__main__":
    main()