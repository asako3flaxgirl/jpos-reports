# jpos-reports
Reports for JP Open Studios.

# How to run

1. Make sure [Python 3.12](https://www.python.org/downloads/release/python-3123/) is installed on your computer
2. Open `registration_report.py` and set the following parameters near the top of the file:

```
WOOCOMMERCE_PATH = "" # location of Woocommerce export, must be .csv
WPFORMS_PATH = ""     # location of WPForms export, must be .xlsx
START_DATE = ""       # the beginning date of the date range in YYYY-MM-DD format
END_DATE = ""         # the last date of the date range in YYYY-MM-DD format
```

3. Run the following commands:

```
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3.12 -m registration_report
```
