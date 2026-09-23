import asyncio
import logging
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright


# -----------------------------
# Configuration
# -----------------------------
SEARCH_URL = (
    "https://www.naukri.com/python-developer-jobs-in-chennai"
    "?k=python%20developer&l=chennai&experience=0"
)

EXCEL_FILE = Path("naukri_jobs.xlsx")
LOG_FILE = Path("scraper.log")

# -----------------------------
# Logging
# -----------------------------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# -----------------------------
# Helpers
# -----------------------------
async def get_text(locator):
    try:
        if await locator.count() > 0:
            text = await locator.first.inner_text()
            return text.strip()
    except Exception:
        pass
    return ""


async def scrape_jobs():
    jobs = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)

        page = await browser.new_page(
            viewport={"width": 1366, "height": 768}
        )

        try:
            logging.info("Opening Naukri search page")

            await page.goto(
                SEARCH_URL,
                wait_until="domcontentloaded",
                timeout=60000,
            )

            await page.wait_for_timeout(5000)

            # Scroll to load more job cards
            for _ in range(4):
                await page.mouse.wheel(0, 1200)
                await page.wait_for_timeout(1500)

            # Current Naukri job-card selector
            cards = page.locator("div.srp-jobtuple-wrapper")

            count = await cards.count()

            logging.info("Job cards found: %s", count)

            print(f"\nJob cards found: {count}\n")

            for i in range(count):
                try:
                    card = cards.nth(i)

                    title = await get_text(card.locator("a.title"))
                    company = await get_text(card.locator("a.comp-name"))
                    location = await get_text(card.locator("span.locWdth"))
                    experience = await get_text(card.locator("span.expwdth"))
                    skills = await get_text(card.locator("ul.tags-gt li"))
                    posted_date = await get_text(
                        card.locator("span.job-post-day")
                    )

                    job_url = ""

                    try:
                        job_url = await card.locator(
                            "a.title"
                        ).first.get_attribute("href")
                    except Exception:
                        pass

                    # Fallback selectors
                    if not title:
                        title = await get_text(
                            card.locator("[class*='title']")
                        )

                    if not company:
                        company = await get_text(
                            card.locator("[class*='comp']")
                        )

                    if not location:
                        location = await get_text(
                            card.locator("[class*='loc']")
                        )

                    if not experience:
                        experience = await get_text(
                            card.locator("[class*='exp']")
                        )

                    # Skills
                    if not skills:
                        try:
                            skill_items = card.locator("ul.tags-gt li")
                            skill_count = await skill_items.count()

                            skill_list = []

                            for j in range(skill_count):
                                skill_text = (
                                    await skill_items.nth(j).inner_text()
                                )
                                if skill_text:
                                    skill_list.append(skill_text.strip())

                            skills = ", ".join(skill_list)
                        except Exception:
                            skills = ""

                    if title and job_url:
                        jobs.append(
                            {
                                "Job Title": title,
                                "Company": company,
                                "Location": location,
                                "Experience": experience,
                                "Skills": skills,
                                "Posted Date": posted_date,
                                "Job URL": job_url,
                            }
                        )

                except Exception as e:
                    logging.error(
                        "Error scraping job card %s: %s",
                        i,
                        str(e),
                    )

            await browser.close()

        except Exception as e:
            logging.exception("Scraping failed: %s", str(e))
            await browser.close()

    return jobs


# -----------------------------
# Excel handling
# -----------------------------
def save_new_jobs(jobs):
    if not jobs:
        print("No jobs scraped.")
        return

    new_df = pd.DataFrame(jobs)

    # Remove duplicates from current scrape
    new_df = new_df.drop_duplicates(subset=["Job URL"])

    if EXCEL_FILE.exists():
        try:
            old_df = pd.read_excel(EXCEL_FILE)

            if "Job URL" not in old_df.columns:
                old_df["Job URL"] = ""

            existing_urls = set(
                old_df["Job URL"]
                .dropna()
                .astype(str)
                .tolist()
            )

            new_df = new_df[
                ~new_df["Job URL"].astype(str).isin(existing_urls)
            ]

            if len(new_df) > 0:
                final_df = pd.concat(
                    [old_df, new_df],
                    ignore_index=True,
                )

                final_df.to_excel(
                    EXCEL_FILE,
                    index=False,
                )

                print(
                    f"Added {len(new_df)} new job(s)."
                )

            else:
                print("No new jobs found. Excel is unchanged.")

        except Exception as e:
            logging.exception(
                "Error reading/updating Excel: %s",
                str(e),
            )

    else:
        new_df.to_excel(
            EXCEL_FILE,
            index=False,
        )

        print(
            f"Created Excel file with {len(new_df)} job(s)."
        )


# -----------------------------
# Main
# -----------------------------
async def main():
    print("Starting Naukri Job Scraper...")
    logging.info("Scraper started")

    jobs = await scrape_jobs()

    print(f"Scraped jobs: {len(jobs)}")

    save_new_jobs(jobs)

    logging.info("Scraper finished")
    print("Done.")


if __name__ == "__main__":
    asyncio.run(main())