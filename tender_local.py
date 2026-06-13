from playwright.sync_api import sync_playwright
import pandas as pd
from urllib.parse import urljoin

SITES = [
    {
        "name": "Central eProcure",
        "url": "https://eprocure.gov.in/eprocure/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Haryana eTenders",
        "url": "https://etenders.hry.nic.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "PMGSY",
        "url": "https://pmgsytenders.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Defence eProcurement",
        "url": "https://defproc.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "eTenders India",
        "url": "https://etenders.gov.in/eprocure/app?page=FrontEndLatestActiveTenders&service=page"
    }
]

TARGET_STATES = [
    "himachal",
    "himachal pradesh",
    "haryana",
    "punjab",
    "uttarakhand",
    "uttrakhand",
    "rajasthan",
    "jammu",
    "jammu and kashmir",
    "j&k"
]

INCLUDE = [
    "rcc road",
    "road construction",
    "construction of road",
    "building construction",
    "construction of building",
    "civil work",
    "civil works",
    "infrastructure",
    "infrastructure development",
    "prefab",
    "prefabricated",
    "peb"
]

EXCLUDE = [
    "cc road",
    "repair",
    "repairs",
    "maintenance",
    "renovation",
    "alteration",
    "electrical",
    "hvac",
    "painting",
    "whitewash",
    "supply",
    "labour",
    "labor",
    "manpower",
    "consultancy",
    "housekeeping",
    "security"
]

MIN_VALUE = 10000000
MIN_SCORE = 15
BATCH_SIZE = 3

def clean(text):
    return " ".join(text.split())

def parse_value(value_text):
    value_text = value_text.replace(",", "").replace("₹", "").strip()

    if value_text.upper() == "NA" or value_text == "":
        return None

    try:
        return int(float(value_text))
    except:
        return None

def calculate_score(title, org, value):
    text = (title + " " + org).lower()
    score = 0

    if "rcc road" in text:
        score += 10
    if "road construction" in text or "construction of road" in text:
        score += 8
    if "building construction" in text or "construction of building" in text:
        score += 10
    if "civil work" in text or "civil works" in text:
        score += 5
    if "infrastructure" in text:
        score += 5
    if "prefab" in text or "prefabricated" in text or "peb" in text:
        score += 10
    if value is not None and value >= MIN_VALUE:
        score += 10

    return score

def state_match(title, org):
    if not TARGET_STATES:
        return True

    text = (title + " " + org).lower()
    return any(state in text for state in TARGET_STATES)

def is_relevant(title, org, value):
    text = (title + " " + org).lower()

    if any(word in text for word in EXCLUDE):
        return False

    if not any(word in text for word in INCLUDE):
        return False

    if not state_match(title, org):
        return False

    if value is None:
        return False

    if value < MIN_VALUE:
        return False

    if calculate_score(title, org, value) < MIN_SCORE:
        return False

    return True

def get_table_rows(page):
    selectors = [
        "table#table tr",
        "table.list_table tr",
        "tr"
    ]

    for selector in selectors:
        try:
            rows = page.locator(selector).all()
            if len(rows) > 1:
                return rows
        except:
            pass

    return []

def extract_current_page(page, site_name):
    data = []
    rows = get_table_rows(page)

    if not rows:
        print(f"{site_name}: No rows found on this page.")
        return data

    for row in rows:
        try:
            cells = row.locator("td").all()
        except:
            continue

        if len(cells) < 7:
            continue

        try:
            s_no = clean(cells[0].inner_text())
            published = clean(cells[1].inner_text())
            closing = clean(cells[2].inner_text())
            opening = clean(cells[3].inner_text())
            title_ref = clean(cells[4].inner_text())
            org = clean(cells[5].inner_text())
            value_text = clean(cells[6].inner_text())
        except:
            continue

        if not s_no.replace(".", "").isdigit():
            continue

        value = parse_value(value_text)

        if not is_relevant(title_ref, org, value):
            continue

        link = page.url

        try:
            anchor = cells[4].locator("a").first
            href = anchor.get_attribute("href")
            if href:
                link = urljoin(page.url, href)
        except:
            pass

        score = calculate_score(title_ref, org, value)

        data.append({
            "Portal": site_name,
            "Score": score,
            "S.No": s_no,
            "Published Date": published,
            "Closing Date": closing,
            "Opening Date": opening,
            "Tender Title / Ref / ID": title_ref,
            "Organisation": org,
            "Tender Value": value_text,
            "Direct Link": link
        })

    return data

def has_next_page(page):
    next_texts = ["Next >", "Next", ">"]

    for text in next_texts:
        try:
            button = page.locator(f"text={text}").first
            if button.is_visible():
                return True
        except:
            pass

    return False

def go_to_next_page(page):
    next_texts = ["Next >", "Next", ">"]

    for text in next_texts:
        try:
            button = page.locator(f"text={text}").first
            if button.is_visible():
                button.click()
                page.wait_for_timeout(2500)
                return True
        except:
            pass

    return False

def read_all_pages(page, site_name):
    all_site_tenders = []
    page_number = 1

    while True:
        print(f"{site_name}: Reading page {page_number}...")

        current_page_tenders = extract_current_page(page, site_name)
        all_site_tenders.extend(current_page_tenders)

        print(f"{site_name}: Page {page_number}: {len(current_page_tenders)} matching tenders found.")

        if not has_next_page(page):
            break

        if not go_to_next_page(page):
            break

        page_number += 1

    return all_site_tenders

def open_batch(context, batch):
    pages = []

    for site in batch:
        print(f"Opening portal: {site['name']}")
        page = context.new_page()
        page.goto(site["url"], wait_until="domcontentloaded")
        pages.append({
            "site": site,
            "page": page
        })

    return pages

def process_batch(context, batch):
    opened_pages = open_batch(context, batch)

    print("\nComplete captcha manually in all opened tabs.")
    print("Click Search on each portal.")
    print("When tender lists are visible in all tabs, return to terminal and press Enter.")

    input("\nPress Enter after all portals in this batch are loaded...")

    batch_tenders = []

    for item in opened_pages:
        site_name = item["site"]["name"]
        page = item["page"]

        try:
            tenders = read_all_pages(page, site_name)
            batch_tenders.extend(tenders)
            print(f"{site_name}: Total matching tenders: {len(tenders)}")
        except Exception as error:
            print(f"{site_name}: Failed to read data. Error: {error}")

    for item in opened_pages:
        try:
            item["page"].close()
        except:
            pass

    return batch_tenders

def save_excel(tenders):
    print("\nTotal rows before Excel:", len(tenders))

    if not tenders:
        print("No matching tenders found.")
        return

    df = pd.DataFrame(tenders)
    df = df.drop_duplicates(subset=["Tender Title / Ref / ID", "Organisation"], keep="first")
    df = df.sort_values(by=["Score", "Tender Value"], ascending=False)
    df.to_excel("all_filtered_tenders.xlsx", index=False)

    print("Excel file created successfully.")
    print("Output file: all_filtered_tenders.xlsx")

def main():
    all_tenders = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()

        for index in range(0, len(SITES), BATCH_SIZE):
            batch = SITES[index:index + BATCH_SIZE]
            print(f"\nStarting batch {index // BATCH_SIZE + 1}")
            batch_tenders = process_batch(context, batch)
            all_tenders.extend(batch_tenders)

        save_excel(all_tenders)

        print(f"\nCompleted! {len(all_tenders)} total matching tenders saved.")
        input("\nPress Enter to close the browser...")
        browser.close()

if __name__ == "__main__":
    main()