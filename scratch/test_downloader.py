import sys
import os
sys.path.append(os.getcwd())
from src import downloader

# Test with DMC and 2025 (or 2024 if 2025 doesn't exist yet)
tickers = ["DMC"]
years = [2025, 2024]

print(f"Testing downloader with tickers={tickers} and years={years}")
downloaded = downloader.run(tickers=tickers, years=years)
print(f"Downloaded files: {downloaded}")