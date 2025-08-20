import numpy as np
import pandas as pd
import os
from datetime import datetime

# Time blocks you want to aggregate by
time_blocks = [
    ('10:30', '10:59'),
    ('11:00', '11:59'),
    ('12:00', '12:30'),
    ('14:30', '14:59'),
    ('15:00', '15:59'),
    ('16:00', '16:59'),
    ('17:00', '18:00'),
    ('21:00', '21:59'),
    ('22:00', '22:59'),
    ('23:00', '23:30')
]

# Set the folder path
input_folder = 'C:/Users/bwlau/Desktop/FCPO/外盘期货3品种/国内（北京时间）/马棕油/1分钟'
output_folder = 'C:/Users/bwlau/Desktop/FCPO/外盘期货3品种/国内（北京时间）/马棕油/1 Hour'

# Loop through all relevant CSV files
for filename in os.listdir(input_folder):
    if filename.endswith('.csv') and '马棕油' in filename:
        filepath = os.path.join(input_folder, filename)

        # Extract datenum for Excel sheet name (optional, fallback if needed)
        datenum = filename.split('（')[1].split('）')[0]

        po = pd.read_csv(filepath, skiprows=1, usecols=list(range(0, 7)),
                         encoding='ISO-8859-1', names=["Date", "Open", "High", "Low", "Close", "OI", "Volume"])
        po = po[(po['Open'] != 0) & (po['Close'] != 0)].fillna(0)

        data = po.reset_index(drop=True).rename(columns={'Close': 'Last'})
        data['Date'] = pd.to_datetime(data['Date'])

        Date_reference = pd.to_datetime(data['Date']).dt.date.drop_duplicates().tolist()
        compressed_data = []

        for date in Date_reference:
            for start_str, end_str in time_blocks:
                start_time = pd.to_datetime(start_str, format='%H:%M').time()
                end_time = pd.to_datetime(end_str, format='%H:%M').time()

                start_datetime = datetime.combine(date, start_time)
                end_datetime = datetime.combine(date, end_time)

                filtered_data = data[(data['Date'] >= start_datetime) & (data['Date'] <= end_datetime)]

                if not filtered_data.empty:
                    open_price = filtered_data.iloc[0]['Open']
                    close_price = filtered_data.iloc[-1]['Last']
                    high_price = filtered_data[['Open', 'High', 'Low', 'Last']].max().max()
                    low_price = filtered_data[['Open', 'High', 'Low', 'Last']].min().min()

                    compressed_data.append({
                        'Date': end_datetime.strftime('%Y-%m-%d %H:%M'),
                        'Start': start_datetime,
                        'End': end_datetime,
                        'Open': open_price,
                        'High': high_price,
                        'Low': low_price,
                        'Close': close_price
                    })

        compressed_df = pd.DataFrame(compressed_data)

        # Output file with the same name (but .xlsx)
        output_file = os.path.join(output_folder, filename.replace('.csv', '.xlsx'))
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
            compressed_df.to_excel(writer, sheet_name='Compressed', index=False)

        print(f'{filename}: Data saved to {output_file}')

#---------------------------------------------------------------------------
#%%
import pandas as pd
import os
from datetime import datetime, timedelta

# Folder containing the Excel files
input_folder = 'C:/Users/bwlau/Desktop/FCPO/外盘期货3品种/国内（北京时间）/马棕油/1 Hour'

# Start and end date of the full range
start_date = pd.to_datetime('2014-09-1')
end_date = pd.to_datetime('2025-08-12')

# List all Excel files in the folder
all_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') and '马棕油' in f]

# Build all periods (16th of each month to 15th of the next)
periods = []
current_start = start_date

while current_start < end_date:
    next_month = current_start + pd.DateOffset(months=1)
    current_end = pd.to_datetime(f"{next_month.year}-{next_month.month:02d}-15")
    if current_end > end_date:
        current_end = end_date

    periods.append((current_start, current_end))
    current_start = current_end + timedelta(days=1)

# Collect extracted data
combined_df = pd.DataFrame()

for period_start, period_end in periods:
    # Determine the contract month based on cutoff day
    if period_start.day <= 15:
        file_month = (period_start + pd.DateOffset(months=2)).month
        file_year = (period_start + pd.DateOffset(months=2)).year
    else:
        file_month = (period_start + pd.DateOffset(months=3)).month
        file_year = (period_start + pd.DateOffset(months=3)).year

    month_code = f"{file_month:02d}"
    year_code = str(file_year)

    # Try to find a matching file (e.g., "马棕油03（2025" in filename)
    matching_files = [f for f in all_files if f"马棕油{month_code}（{year_code}" in f]

    if not matching_files:
        print(f"❌ No file found for contract 马棕油{month_code}（{year_code}）")
        continue

    filename = matching_files[0]
    filepath = os.path.join(input_folder, filename)

    try:
        df = pd.read_excel(filepath, sheet_name='Compressed')
        df['Date'] = pd.to_datetime(df['Date'])

        # Filter data within the actual date range
        filtered = df[(df['Date'] >= period_start) & (df['Date'] <= period_end)]

        if not filtered.empty:
            filtered['ContractMonth'] = f"马棕油{month_code}"  # ← Column H
            combined_df = pd.concat([combined_df, filtered], ignore_index=True)
            print(f"✅ {filename} → {period_start.date()} to {period_end.date()}, rows: {len(filtered)}")
        else:
            print(f"⚠️ {filename} has no data in range {period_start.date()} to {period_end.date()}")

    except Exception as e:
        print(f"❌ Error reading {filename}: {e}")

# Save final output
output_path = os.path.join(input_folder, 'combined_extracted_data.xlsx')
combined_df.to_excel(output_path, index=False)

print(f"\n✅ All extracted data saved to: {output_path}")
