# minute-to-hourly-etl
Data clearning project: converting minute-level data to hourly for analysis

🛢️ Palm Oil Futures Intraday Data Aggregation & Extraction

This project provides a set of Python scripts to aggregate and extract intraday market data (1-minute frequency) for Malaysian Palm Oil Futures (马棕油) into defined hourly blocks, and then organize that data into contract-based monthly periods for further analysis.

📦 Features

Intraday Aggregation:

Converts 1-minute data into predefined time blocks (e.g., 10:30–10:59, 11:00–11:59, etc.).

Calculates Open, High, Low, Close prices for each block.

Saves aggregated data into .xlsx files, one per contract month.

Period-Based Extraction:

Extracts data within a rolling window (16th of each month to the 15th of the next).

Assigns data to the correct contract month based on exchange rollover conventions.

Combines all valid data into a single consolidated Excel sheet.

📁 Folder Structure
.
├── 1分钟/                          # Input folder: Raw 1-minute data
│   └── 马棕油*.csv                # CSV files with 1-minute market data
├── 1 Hour/                        # Output folder: Aggregated hourly data
│   └── 马棕油*.xlsx              # Hourly data in Excel format
└── combined_extracted_data.xlsx  # Final consolidated dataset

🕒 Time Blocks Used for Aggregation

The script aggregates data into the following time blocks:

10:30–10:59

11:00–11:59

12:00–12:30

14:30–14:59

15:00–15:59

16:00–16:59

17:00–18:00

21:00–21:59

22:00–22:59

23:00–23:30

These blocks are designed to match active trading sessions in local (China/Beijing) time.

🧠 Logic for Contract Mapping

Each 16th-to-15th period is mapped to a contract month that is 2 months ahead, unless the start day is after the 15th, in which case it's mapped 3 months ahead. This mimics the contract rolling logic used in futures trading.

For example:

Period Start	Mapped Contract
2025-06-16	马棕油09
2025-07-01	马棕油10
⚙️ Requirements

Python 3.x

pandas

openpyxl

numpy

Install with:

pip install pandas openpyxl numpy

📌 How to Use

Place all raw .csv 1-minute files in the 1分钟/ folder.

Run the first script to convert them into hourly time blocks (1 Hour/*.xlsx).

Run the second script to extract data for each rolling period and generate a combined Excel output.

📝 Notes

Filenames must contain the contract month (e.g., 马棕油03（2025）) to ensure proper matching.

Assumes all timestamps are already in Beijing time.

You can adjust time_blocks or date ranges as needed.

📊 Output

The final combined Excel file (combined_extracted_data.xlsx) contains:

Date	Start	End	Open	High	Low	Close	ContractMonth
2025-07-16 10:59	2025-07-16 10:30	2025-07-16 10:59	...	...	...	...	马棕油10
🔒 License

This project is intended for personal or research use. Please check licensing and data source terms before public or commercial use.
