import asyncio
from pathlib import Path
import pandas as pd
from playwright.async_api import async_playwright


async def get_text_content_by_xpath(page, xpath) -> str:
    element = page.locator(xpath)
    return await element.text_content()


async def get_info(page) -> tuple[str, str]:
    # xpaths for user and receipt IDs
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


async def get_records(start_url) -> list[dict[str, str]]:
    records = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        await page.goto(start_url)

        # Pause for user interactions
        input("Navigate to first student submission. Press Enter when done...")

        # Downloading ID mappings
        print("ID mapping download: START")
        while True:
            uid, rcp = await get_info(page)
            records.append({"uid": uid, "rcp": rcp})
            button_next = page.locator('[aria-label="Next Student"]')
            if await next_clickable(button_next):
                await button_next.click()
                await page.wait_for_timeout(1000)  # Wait for navigation to settle
            else:
                break
        print("ID mapping download: END")
        print(f"Retrieved {len(records)} records")

        await browser.close()

    return records


def save_records(records: list[dict[str, str]], path: Path) -> Path:
    df = pd.DataFrame(records)
    df.to_excel(path)
    print(f"File saved to: {path}")
    return path


async def main():
    # TODO: turn to argparse
    # Arguments
    d_out = Path("1b-ids_mapping")
    assert d_out.exists()
    stamp = pd.Timestamp.now().isoformat(timespec="seconds").replace(":", "")
    p_out = d_out / f"id_mapping-{stamp}.xlsx"

    # Execution
    records = await get_records(
        "https://www.learn.ed.ac.uk/auth-saml/saml/login?apId=_175_1"
    )
    save_records(records, p_out)


if __name__ == "__main__":
    asyncio.run(main())
