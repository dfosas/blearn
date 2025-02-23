# coding=utf-8
import logging
from pathlib import Path

import fire
import pandas as pd
from playwright.async_api import async_playwright

from blearn.utils import setup_logger


async def get_text_content_by_xpath(page, xpath) -> str:
    element = page.locator(xpath)
    return await element.text_content()


async def get_info(page) -> tuple[str, str]:
    # XPaths for user and receipt IDs
    xpath_uid = """//*[@id="main-content"]/div[7]/div/div/div/div/bb-flexible-attempt-grading-ui/section/header/div/header/div/div/div[1]/div/div[2]/div/h1/div/div/bdi"""
    uid = await get_text_content_by_xpath(page, xpath_uid)
    xpath_rcp = """//*[@id="main-content"]/div[7]/div/div/div/div/bb-flexible-attempt-grading-ui/section/header/div/header/div/div/div[2]/div[1]/div[1]/span"""
    rcp = await get_text_content_by_xpath(page, xpath_rcp)
    return uid, rcp


async def next_clickable(element) -> bool:
    is_disabled = await element.get_attribute("aria-disabled")
    if is_disabled == "true":
        return False
    elif is_disabled == "false":
        return True
    else:
        raise ValueError(f"Exception with {is_disabled=}")


async def get_records(start_url: str, wait: bool = False) -> list[dict[str, str]]:
    """
    Get records for anonymous submissions.

    Parameters
    ----------
    start_url
        The URL to start navigation from for the user.
    wait
        Whether to wait keep the browser opened at the end until user confirmation
        that it can be closed.

    Returns
    -------
    list[dict[str, str]]
        List of dicts with user IDs (uid) and receipt IDs (rcp), one per student-submission.

    """
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        await page.goto(start_url)

        # Pause for user interactions
        input("Navigate to first student submission. Press Enter when done...")

        # Downloading ID mappings
        logging.info("ID mapping download: START")
        records = []
        n = 0
        while True:
            uid, rcp = await get_info(page)
            records.append({"uid": uid, "rcp": rcp})
            n += 1
            logging.info(f"Got '{uid}' ({n} records so far)")
            button_next = page.locator('[aria-label="Next Student"]')
            if await next_clickable(button_next):
                logging.info("There is another submission: moving to the next one")
                await button_next.click()
                await page.wait_for_timeout(1000)  # Wait for navigation to settle
            else:
                logging.info("No more submissions available")
                break
        logging.info("ID mapping download: END")
        logging.info(f"Retrieved {len(records)} records")

        if wait:
            input("Press Enter when ready to close the browser...")
        await browser.close()

    return records


def save_records(records: list[dict[str, str]], path: Path) -> Path:
    logging.info(f"Saving records to: {path}")
    df = pd.DataFrame(records)
    df.to_excel(path)
    return path


async def main(
    start_url: str,
    p_out: Path | None = None,
    p_log: Path | None = None,
    wait: bool = False,
    debug: bool = False,
):
    """
    Web automation tool to extract Learn's anonymous marking IDs.

    Parameters
    ----------
    start_url
        Address to start the process from.
    p_out
        Output file path for the spreadsheet.
    p_log
        Path to log file to write output to.
    wait
        Whether to wait keep the browser opened at the end until user confirmation
        that it can be closed.
    debug
        Activate debug mode for the logger.

    """
    if p_out is None:
        stamp = pd.Timestamp.now().isoformat(timespec="seconds").replace(":", "")
        p_out = Path.cwd() / f"id_mapping-{stamp}.xlsx"
    if p_log is None:
        p_log = p_out.parent / f"blearn-{Path(__file__).stem}.log"
    setup_logger(path=p_log, shout=True, debug=debug)
    logging.info("INI.")
    records = await get_records(start_url=start_url, wait=wait)
    save_records(records, p_out)
    logging.info("END.")


def cli():
    fire.Fire(main)


if __name__ == "__main__":
    fire.Fire(main)
