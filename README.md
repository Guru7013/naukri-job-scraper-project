# Naukri Job Scraper

A Python application using Playwright to scrape Python Developer job listings from Naukri.com.

## Features

- Scrapes job title
- Scrapes company name
- Scrapes location
- Scrapes experience
- Scrapes skills
- Scrapes posted date
- Scrapes job URL
- Saves data to Excel
- Prevents duplicate jobs using Job URL
- Preserves previously scraped jobs
- Adds only new jobs on subsequent runs
- Includes logging and error handling

## Technologies

- Python
- Playwright
- Pandas
- OpenPyXL

## Project Files

- `scraper.py` - Main scraping program
- `naukri_jobs.xlsx` - Scraped job data
- `requirements.txt` - Python dependencies
- `scraper.log` - Application log
- `README.md` - Project documentation

## Installation

Create and activate a virtual environment:

```bash
python -m venv venv