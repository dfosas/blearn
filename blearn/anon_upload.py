# coding=utf-8
import logging
from pathlib import Path

import fire
import pandas as pd
from playwright.async_api import async_playwright

from blearn.utils import setup_logger
from blearn.anon_mapping import get_ids, maybe_go_to_next


async def upload_records(
    records: dict[str, dict[str, str]],
    start_url: str,
    n_max: int | None,
    pause: int | None = None,
    wait_end: bool = False,
) -> None:
    """
    Upload records for anonymous submissions.

    Parameters
    ----------
    records
        Mapping `submission_id` to student, mark and feedback.
    start_url
        The URL to start navigation from for the user.
    n_max
        Limit the number of records to return up to `n_max`. If `None` retrieve all records.
    pause
        Time to pause between submissions in milliseconds. If `None` or 0 then no pause is used.
    wait_end
        Whether to wait keep the browser opened at the end until user confirmation
        that it can be closed.

    """
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        await page.goto(start_url)

        # Pause for user interactions
        input("Navigate to first student submission. Press Enter when done...")

        # Downloading ID mappings
        logging.info("ID mapping download: START")
        n = 0
        while True:
            uid, rcp, _ = await get_ids(page)
            uid = uid.replace("Anonymous Student ", "")
            rcp = rcp.replace("Receipt: ", "")
            n += 1
            logging.info(f"'{uid}': start ({n} records so far)")

            # Open feedback panel (only the first time, Learn remembers...)
            if n == 1:
                element = page.locator('[title="Open feedback panel"]')
                await element.click()

            # Upload marks if available
            if rcp in records:
                logging.info("    case available: uploading...")
                student_id = str(records[rcp]["student_id"])
                assert uid == student_id, f"mismatch: {uid} != {student_id}"
                mark = str(records[rcp]["mark"])
                feedback = records[rcp]["feedback"]

                # A) Write feedback
                # - step 1: write feedback
                await page.fill(selector='[id="bb-editor-textbox"]', value=feedback)
                # - step 2: save feedback
                fdbk_save_sel = '[data-analytics-id="attemptGrading.page.body.overallFeedback.saveButton"]'
                fdbk_save_ele = page.locator(fdbk_save_sel)
                await fdbk_save_ele.click()

                # B) Add mark
                grade_sel = '[analytics-id="attemptGrading.students.userWithAttemptHeader.attemptGradePill.gradePill.gradableContent.input"]'
                grade_ele = page.locator(grade_sel)
                await grade_ele.click()
                await page.fill(selector=grade_sel, value=mark)
                await page.keyboard.press("Enter")
            else:
                logging.info("    case unavailable: moving to next...")
            logging.info(f"'{uid}': end")

            next_avail = await maybe_go_to_next(page=page, pause=pause)
            if not next_avail or (n_max is not None and n >= n_max):
                break

        logging.info("ID mapping download: END")
        logging.info(f"Retrieved {len(records)} records")

        if wait_end:
            input("Press Enter when ready to close the browser...")
        await browser.close()


async def main(
    p_xlsx: Path,
    start_url: str,
    p_log: Path | None = None,
    n_max: int | None = None,
    pause: int | None = None,
    wait_end: bool = False,
    debug: bool = False,
):
    """
    Web automation tool to extract Learn's anonymous marking IDs.

    Parameters
    ----------
    p_xlsx
        Path to spreadsheet with student_id, submission_id, mark, feedback.
    start_url
        Address to start the process from.
    p_log
        Path to log file to write output to.
    n_max
        Limit the number of records to return up to `n_max`. If `None` retrieve all records.
    pause
        Time to pause between submissions in milliseconds. If `None` or 0 then no pause is used.
    wait_end
        Whether to wait keep the browser opened at the end until user confirmation
        that it can be closed.
    debug
        Activate debug mode for the logger.

    """
    if p_log is None:
        p_log = Path.cwd() / f"blearn-{Path(__file__).stem}.log"
    setup_logger(path=p_log, shout=True, debug=debug)
    logging.info("INI.")
    records = pd.read_excel(p_xlsx).set_index("submission_id").to_dict(orient="index")
    logging.info(f"Starting upload of {len(records)} records...")
    await upload_records(
        records=records,
        start_url=start_url,
        n_max=n_max,
        pause=pause,
        wait_end=wait_end,
    )
    logging.info("END.")


def cli():
    fire.Fire(main)


if __name__ == "__main__":
    fire.Fire(main)
