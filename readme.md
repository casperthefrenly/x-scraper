# X (Twitter) Engagement Scraper & Reporter

An automated Python tool designed to scrape engagement metrics (Likes, Views, Retweets, and Comments) from X.com and generate a nicely formatted Excel report.

## Features
Asynchronous Scraping: Utilizes `Playwright` and `asyncio` to process multiple URLs concurrently for maximum efficiency.
Automated Data Extraction: Uses intelligent selectors and Regex to find metrics even when UI elements vary.
Formated Excel Output: Automatically formats the resulting spreadsheet with specific fonts (Arial), cell borders, center-alignment, and auto-adjusted column widths.
Timezone Aware: Converts tweet timestamps to `CET` (customizable) for accurate reporting.

## Installation & Setup

git clone https://github.com/casperthefrenly/x-scraper.git
cd x-scraper
pip install -r requirements.txt
playwright install chromium

## How to Use

Open `tweets_input.xlsx`.
Put your X.com links in **Column A**.
Run the Script.
View Results in `scraped_tweets_output.xlsx`.
