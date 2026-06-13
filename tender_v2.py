import asyncio
# import aiohttp  # only needed when AI classifier is enabled
import json
import re
import pandas as pd
from urllib.parse import urljoin
from playwright.async_api import async_playwright

# ── SITES ─────────────────────────────────────────────────────────────────────

SITES = [
    {
        "name": "Central eProcure",
        "file": "central_eprocure.xlsx",
        "state_portal": False,
        "state_keywords": [
            "himachal", "himachal pradesh", "haryana", "punjab",
            "uttarakhand", "uttrakhand", "rajasthan"
        ],
        "slow_pagination": False,
        "timeout_ms": 45000,
        "url": "https://eprocure.gov.in/eprocure/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Haryana eTenders",
        "file": "haryana_etenders.xlsx",
        "state_portal": True,
        "state_keywords": [],
        "slow_pagination": False,
        "timeout_ms": 45000,
        "url": "https://etenders.hry.nic.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Punjab eProcurement",
        "file": "punjab_eprocurement.xlsx",
        "state_portal": True,
        "state_keywords": [],
        "slow_pagination": False,
        "timeout_ms": 45000,
        "url": "https://eproc.punjab.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Rajasthan eProcurement",
        "file": "rajasthan_eprocurement.xlsx",
        "state_portal": True,
        "state_keywords": [],
        "slow_pagination": False,
        "timeout_ms": 45000,
        "url": "https://eproc.rajasthan.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Himachal eTenders",
        "file": "himachal_etenders.xlsx",
        "state_portal": True,
        "state_keywords": [],
        "slow_pagination": False,
        "timeout_ms": 45000,
        "url": "https://hptenders.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "Uttarakhand eTenders",
        "file": "uttarakhand_etenders.xlsx",
        "state_portal": True,
        "state_keywords": [],
        "slow_pagination": False,
        "timeout_ms": 45000,
        "url": "https://uktenders.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "JK eTenders",
        "file": "jammu_etenders.xlsx",
        "state_portal": False,
        "state_keywords": ["jammu"],
        "slow_pagination": True,
        "timeout_ms": 90000,
        "url": "https://jktenders.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
    {
        "name": "PMGSY",
        "file": "pmgsy.xlsx",
        "state_portal": False,
        "state_keywords": [
            "himachal", "himachal pradesh", "haryana", "punjab",
            "uttarakhand", "uttrakhand", "rajasthan", "jammu"
        ],
        "slow_pagination": True,
        "timeout_ms": 90000,
        "url": "https://pmgsytenders.gov.in/nicgep/app?page=FrontEndLatestActiveTenders&service=page"
    },
]

# ── KEYWORDS ──────────────────────────────────────────────────────────────────

SPECIAL_KEYWORDS = [
    "prefab", "prefabricated", "pre fabricated",
    "peb", "pre engineered building", "pre-engineered building",
    "pre engineered steel", "steel structure building",
    "modular building", "portable cabin",
    "puf panel", "puff panel", "sandwich panel",
]


BEST_KEYWORDS = [
    "building construction", "construction of building",
    "construction of rcc building", "construction of g+",
    "residential building", "commercial building",
    "institutional building", "industrial building",
    "multi storey building", "multistorey building",
    "school building", "college building", "hospital building",
    "hostel building", "office building", "administrative building",
    "government building", "community building",
    "warehouse building", "factory building",
    "staff quarter", "type iv quarter", "type iii quarter", "type ii quarter",
]

STRONG_KEYWORDS = [
    "construction of school", "construction of college",
    "construction of hostel", "construction of hospital",
    "construction of dispensary", "construction of office",
    "construction of warehouse", "construction of factory",
    "construction of workshop", "construction of community hall",
    "construction of quarter", "construction of residence",
    "construction of dormitory", "construction of auditorium",
    "construction of library", "construction of laboratory",
    "construction of shed", "industrial shed",
    "warehouse shed", "factory shed",
    "godown construction", "construction of godown",
]

GOOD_KEYWORDS = [
    "civil construction", "rcc work", "rcc construction", "structural work",
]

WEAK_KEYWORDS = []

NEGATIVE_WORDS = [
    "repair", "maintenance", "painting", "electrical", "amc",
    "renovation", "retrofitting", "white washing", "whitewashing",
    "plumbing", "flooring", "false ceiling", "interior", "furniture",
    "road", "highway", "bridge", "culvert", "drain", "drainage",
    "sewer", "sewerage", "pipeline", "water supply", "irrigation",
    "canal", "dam", "hydro", "borewell", "boundary wall",
    "compound wall", "retaining wall", "fencing", "solar", "cctv",
    "lift", "elevator", "fire fighting", "firefighting", "hvac",
    "air conditioning", "landscaping", "street light",
    "supply of", "procurement of", "purchase of",
]

HARD_EXCLUDE = [
    " road ", "road work", "road construction", "highway", "flyover",
    "overbridge", "underpass", "bridge construction", "tunnel",
    "irrigation", "canal", "dam", "sewage", "water supply",
    "water treatment", "borewell", "boring work", "electrical work",
    "electrification", "substation", "transformer",
    "solar panel", "solar power", "street light",
    "supply and installation", "supply & installation",
    "laying of", "laying and jointing", "pipe laying",
]

CATEGORY_KEYWORDS = {
    "Construction": BEST_KEYWORDS + STRONG_KEYWORDS + GOOD_KEYWORDS,
    "Prefab / PEB": SPECIAL_KEYWORDS,
}

ALL_RELEVANT_KEYWORDS = SPECIAL_KEYWORDS + BEST_KEYWORDS + STRONG_KEYWORDS + GOOD_KEYWORDS

# ── ABBREVIATION MAP ──────────────────────────────────────────────────────────
# Government portals use heavy abbreviations — expand before keyword matching

ABBREVIATIONS = [
    # Construction — order matters: longer patterns first
    (r"c/o",                        "construction of"),
    (r"c\.o\.",                   "construction of"),
    (r"constn\.",                  "construction"),
    (r"constn",                     "construction"),
    (r"constrn\.",                 "construction"),
    (r"constrn",                    "construction"),
    (r"const\.",                   "construction"),

    # Building
    (r"bldg\.",                    "building"),
    (r"bldg",                       "building"),
    (r"blk\.",                     "block"),

    # Quarters / Residential
    (r"qtrs\.",                    "quarters"),
    (r"qtrs",                       "quarters"),
    (r"qtr\.",                     "quarter"),
    (r"qtr",                        "quarter"),
    (r"residl\.",                  "residential"),
    (r"residl",                     "residential"),

    # Institutional
    (r"hosp\.",                    "hospital"),
    (r"schl\.",                    "school"),
    (r"coll\.",                    "college"),

    # Shed / Godown
    (r"godn\.",                    "godown"),
    (r"godn",                       "godown"),

    # Prefab / PEB
    (r"p\.e\.b\.",               "peb"),
    (r"pre-eng\.",                 "pre engineered building"),
    (r"pre\s+eng\.",              "pre engineered building"),

    # Noise prefixes — remove
    (r"sh:\s*",                    ""),
    (r"nh:\s*",                    ""),
    (r"nit\s+for\s+",            ""),
    (r"providing\s+and\s+fixing", "construction of"),
    (r"providing\s+&\s+fixing",   "construction of"),
    (r"erection\s+of",            "construction of"),
    (r"erection\s+&",             "construction of"),
    (r"ph-?[0-9]+",                ""),
    (r"phase-?[0-9]+",             ""),
]

# ── NEGATIVE ORG NAMES ────────────────────────────────────────────────────────
# If the organisation itself is clearly not construction-related, reject

ORG_EXCLUDE_KEYWORDS = [
    "irrigation",
    "power corporation",
    "power transmission",
    "electricity board",
    "electricity department",
    "powercom",
    "vidyut",
    "jal board",
    "jal nigam",
    "jal shakti",
    "water board",
    "water supply",
    "sewerage board",
    "highway",
    "road construction",
    "national highway",
    "nhai",
    "forest department",
    "horticulture",
    "agriculture",
    "fisheries",
    "animal husbandry",
    "sericulture",
    "telecom",
    "bsnl",
    "railways",
    "metro rail",
    "airport authority",
    "oil and natural gas",
    "ongc",
    "coal india",
    "mining",
]

# Minimum score a tender must have to pass (rejects near-zero matches)
MIN_SCORE_THRESHOLD = 10

# ── CONFIG ────────────────────────────────────────────────────────────────────

MIN_VALUE         = 10000000   # ₹1 Cr
MAX_VALUE         = 200000000  # ₹20 Cr
BEST_SCORE        = 50
REVIEW_SCORE      = 30
FAST_WAIT_MS      = 400
SLOW_WAIT_MS      = 1000
MAX_PAGES_PER_PORTAL = 1000
MAX_OPEN_RETRIES  = 3

# Claude AI classifier settings
AI_BATCH_SIZE     = 20         # titles sent per API call (keeps cost low)
AI_MODEL          = "claude-haiku-4-5-20251001"   # cheapest + fast
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_API_KEY = "YOUR_API_KEY_HERE"           # ← replace with your key

# ── HELPERS ───────────────────────────────────────────────────────────────────

def clean(text):
    return " ".join(str(text).split())


def expand_abbreviations(text):
    """Expand common government tender abbreviations before keyword matching."""
    text = text.lower()
    for pattern, replacement in ABBREVIATIONS:
        text = re.sub(pattern, replacement, text)
    text = " ".join(text.split())
    return text


def parse_value(value_text):
    value_text = str(value_text).replace(",", "").replace("₹", "").strip()
    if value_text.upper() == "NA" or value_text == "":
        return None
    try:
        return int(float(value_text))
    except:
        return None


def value_sort_number(value_text):
    v = parse_value(value_text)
    return v if v is not None else 0


def state_match(title, org, site):
    if site["state_portal"]:
        return True
    text = (title + " " + org).lower()
    return any(s in text for s in site["state_keywords"])


def is_relevant(text):
    """Check against expanded text."""
    expanded = expand_abbreviations(text)
    return any(w in expanded for w in ALL_RELEVANT_KEYWORDS)


def is_bad_org(org):
    """Returns True if the organisation is clearly not a building construction dept."""
    org_lower = org.lower()
    return any(w in org_lower for w in ORG_EXCLUDE_KEYWORDS)


def calculate_score(title, org, value):
    # Use expanded text for better abbreviation handling
    expanded = expand_abbreviations(title + " " + org)
    score = 0
    for w in SPECIAL_KEYWORDS:
        if w in expanded: score += 25
    for w in BEST_KEYWORDS:
        if w in expanded: score += 30
    for w in STRONG_KEYWORDS:
        if w in expanded: score += 22
    for w in GOOD_KEYWORDS:
        if w in expanded: score += 12
    for w in NEGATIVE_WORDS:
        if w in expanded: score -= 15
    if value is not None:
        if value >= 500000000:   score += 20
        elif value >= 100000000: score += 15
        elif value >= MIN_VALUE: score += 10
    return score


def get_tender_type(score):
    if score >= BEST_SCORE:   return "Best Tender"
    if score >= REVIEW_SCORE: return "Review Tender"
    return "Captured Backup"


def is_captured(title, org, value, site):
    """Stage 1 — keyword filter with abbreviation expansion, org filter, score threshold."""
    raw_text = (title + " " + org).lower()

    if value is None or value < MIN_VALUE or value > MAX_VALUE:
        return False

    if not state_match(title, org, site):
        return False

    # Hard-exclude by title keywords
    if any(w in raw_text for w in HARD_EXCLUDE):
        return False

    # Reject if organisation itself is clearly irrelevant
    if is_bad_org(org):
        return False

    # Must match at least one relevant keyword (checked on expanded text)
    if not is_relevant(raw_text):
        return False

    # Score must be above minimum threshold (filters weak/accidental matches)
    score = calculate_score(title, org, value)
    if score < MIN_SCORE_THRESHOLD:
        return False

    return True

# ── STAGE 2: CLAUDE AI CLASSIFIER ─────────────────────────────────────────────

AI_SYSTEM_PROMPT = """You are a tender classification assistant for a construction company in India.
The company ONLY does:
1. Building construction (residential, commercial, institutional, industrial buildings)
2. Prefab buildings
3. Pre-Engineered Buildings (PEB)

You will receive a list of government tender titles. For each, reply YES if it is relevant to the company's work, or NO if it is not.

Relevant = any kind of building construction, prefab structure, PEB, industrial shed, warehouse, factory building, school/hospital/hostel/office building construction.

NOT relevant = roads, bridges, dams, irrigation, pipelines, electrical work, solar, repair/maintenance, supply of goods, IT systems, vehicles, furniture, or anything that is not building construction.

Reply ONLY with a JSON array of "YES" or "NO" in the same order as the input titles.
Example input: ["Construction of Road", "Construction of School Building"]
Example output: ["NO", "YES"]
Do not include any explanation."""


async def ai_classify_batch(session, titles):
    """Send a batch of titles to Claude and get YES/NO for each."""
    prompt = json.dumps(titles)
    payload = {
        "model": AI_MODEL,
        "max_tokens": 500,
        "system": AI_SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": prompt}]
    }
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    try:
        import aiohttp
        async with session.post(ANTHROPIC_API_URL, json=payload, headers=headers, timeout=aiohttp.ClientTimeout(total=30)) as resp:
            if resp.status != 200:
                text = await resp.text()
                print(f"  [AI] API error {resp.status}: {text[:200]}")
                return ["YES"] * len(titles)   # fail open — don't drop tenders on API error

            data = await resp.json()
            raw = data["content"][0]["text"].strip()

            # strip markdown fences if present
            raw = raw.replace("```json", "").replace("```", "").strip()
            results = json.loads(raw)

            if len(results) != len(titles):
                print(f"  [AI] Length mismatch ({len(results)} vs {len(titles)}), failing open.")
                return ["YES"] * len(titles)

            return results

    except Exception as e:
        print(f"  [AI] Exception: {e} — failing open for this batch.")
        return ["YES"] * len(titles)


async def ai_filter_tenders(tenders):
    """
    Stage 2: run all keyword-passed tenders through Claude AI classifier.
    Processes in batches of AI_BATCH_SIZE for efficiency.
    Returns only tenders Claude confirms as relevant.
    """
    if not tenders:
        return []

    print(f"\n[AI Classifier] Checking {len(tenders)} tenders with Claude AI...")

    kept = []
    dropped = 0

    import aiohttp
    async with aiohttp.ClientSession() as session:
        for i in range(0, len(tenders), AI_BATCH_SIZE):
            batch = tenders[i : i + AI_BATCH_SIZE]
            titles = [t["Tender Title / Ref / ID"] for t in batch]

            batch_num = i // AI_BATCH_SIZE + 1
            total_batches = (len(tenders) + AI_BATCH_SIZE - 1) // AI_BATCH_SIZE
            print(f"  [AI] Batch {batch_num}/{total_batches} — {len(titles)} titles...")

            verdicts = await ai_classify_batch(session, titles)

            for tender, verdict in zip(batch, verdicts):
                if str(verdict).strip().upper() == "YES":
                    kept.append(tender)
                else:
                    dropped += 1
                    print(f"    [AI REJECTED] {tender['Tender Title / Ref / ID'][:80]}")

            # small pause between batches to be polite to the API
            if i + AI_BATCH_SIZE < len(tenders):
                await asyncio.sleep(0.5)

    print(f"[AI Classifier] Done. Kept {len(kept)}, rejected {dropped} out of {len(tenders)}.")
    return kept

# ── SCRAPING ──────────────────────────────────────────────────────────────────

async def get_table_rows(page):
    for selector in ["table#table tr", "table.list_table tr", "tr"]:
        try:
            rows = await page.locator(selector).all()
            if len(rows) > 1:
                return rows
        except:
            pass
    return []


async def extract_current_page(page, site):
    data = []
    rows = await get_table_rows(page)

    for row in rows:
        try:
            cells = await row.locator("td").all()
        except:
            continue

        if len(cells) < 7:
            continue

        try:
            s_no       = clean(await cells[0].inner_text())
            published  = clean(await cells[1].inner_text())
            closing    = clean(await cells[2].inner_text())
            opening    = clean(await cells[3].inner_text())
            title_ref  = clean(await cells[4].inner_text())
            org        = clean(await cells[5].inner_text())
            value_text = clean(await cells[6].inner_text())
        except:
            continue

        if not s_no.replace(".", "").isdigit():
            continue

        value = parse_value(value_text)

        # Stage 1: keyword filter
        if not is_captured(title_ref, org, value, site):
            continue

        score       = calculate_score(title_ref, org, value)
        tender_type = get_tender_type(score)
        link        = page.url

        try:
            anchor = cells[4].locator("a").first
            href   = await anchor.get_attribute("href")
            if href:
                link = urljoin(page.url, href)
        except:
            pass

        data.append({
            "Tender Type":            tender_type,
            "Portal":                 site["name"],
            "Score":                  score,
            "S.No":                   s_no,
            "Published Date":         published,
            "Closing Date":           closing,
            "Opening Date":           opening,
            "Tender Title / Ref / ID": title_ref,
            "Organisation":           org,
            "Tender Value":           value_text,
            "Tender Value Number":    value_sort_number(value_text),
            "Direct Link":            link
        })

    return data


async def get_page_signature(page):
    try:
        rows = await get_table_rows(page)
        sigs = []
        for row in rows:
            cells = await row.locator("td").all()
            if len(cells) < 7:
                continue
            s_no = clean(await cells[0].inner_text())
            if not s_no.replace(".", "").isdigit():
                continue
            title = clean(await cells[4].inner_text())
            org   = clean(await cells[5].inner_text())
            sigs.append(s_no + "|" + title[:80] + "|" + org[:40])
            if len(sigs) >= 3:
                break
        return " || ".join(sigs)
    except:
        return ""


async def find_next_button(page):
    for selector in ["a:has-text('Next >')", "a:has-text('Next')", "input[value='Next']", "button:has-text('Next')"]:
        try:
            loc = page.locator(selector)
            if await loc.count() > 0:
                btn = loc.last
                if await btn.is_visible():
                    return btn
        except:
            pass
    try:
        for link in await page.locator("a").all():
            text = clean(await link.inner_text()).lower()
            if text in ["next", "next >", ">"]:
                if await link.is_visible():
                    return link
    except:
        pass
    return None


async def go_to_next_page(page, site):
    wait_ms         = SLOW_WAIT_MS if site["slow_pagination"] else FAST_WAIT_MS
    max_checks      = 45 if site["slow_pagination"] else 24
    max_attempts    = 5  if site["slow_pagination"] else 3
    old_sig         = await get_page_signature(page)

    for _ in range(max_attempts):
        btn = await find_next_button(page)
        if btn is None:
            return False
        try:
            await btn.scroll_into_view_if_needed()
            await page.wait_for_timeout(250)
            await btn.click(force=True)
        except:
            await page.wait_for_timeout(wait_ms)
            continue
        for _ in range(max_checks):
            await page.wait_for_timeout(wait_ms)
            new_sig = await get_page_signature(page)
            if new_sig and new_sig != old_sig:
                return True
        await page.wait_for_timeout(wait_ms)
    return False


def save_portal_excel(site_file, tenders):
    if not tenders:
        return
    df = pd.DataFrame(tenders)
    df = df.drop_duplicates(subset=["Tender Title / Ref / ID", "Organisation"], keep="first")
    df = df.sort_values(by=["Score", "Tender Value Number"], ascending=False)
    df.to_excel(site_file, index=False)


async def is_tender_table_visible(page):
    """Returns True if a real tender table with numbered rows is visible."""
    try:
        rows = await get_table_rows(page)
        for row in rows:
            cells = await row.locator("td").all()
            if len(cells) < 7:
                continue
            s_no = clean(await cells[0].inner_text())
            if s_no.replace(".", "").isdigit():
                return True
    except:
        pass
    return False


async def wait_for_tender_table(page, site, max_wait_seconds=300):
    """
    Waits until the tender table is visible on the page.
    Handles portals that keep redirecting back to captcha/login pages.
    Checks every 3 seconds for up to max_wait_seconds (default 5 minutes).
    """
    site_name = site["name"]
    waited = 0
    check_interval = 3  # seconds

    print(f"{site_name}: Waiting for tender table to appear...")

    while waited < max_wait_seconds:
        if await is_tender_table_visible(page):
            print(f"{site_name}: Tender table detected. Starting scrape.")
            return True

        # Status update every 30 seconds
        if waited > 0 and waited % 30 == 0:
            print(f"{site_name}: Still waiting... ({waited}s). Please solve captcha and click Search.")

        await asyncio.sleep(check_interval)
        waited += check_interval

    print(f"{site_name}: Timed out after {max_wait_seconds}s — tender table never appeared. Skipping.")
    return False


async def read_all_pages(page, site):
    site_name    = site["name"]
    site_tenders = []
    page_number  = 1

    # Wait until the tender table is actually visible before scraping
    if not await wait_for_tender_table(page, site):
        return []

    while True:
        print(f"{site_name}: Reading page {page_number}...")
        current = await extract_current_page(page, site)
        site_tenders.extend(current)
        print(f"{site_name}: Page {page_number}: {len(current)} tenders passed Stage 1.")

        if page_number >= MAX_PAGES_PER_PORTAL:
            print(f"{site_name}: Page limit reached.")
            break

        if not await go_to_next_page(page, site):
            print(f"{site_name}: No next page.")
            break

        page_number += 1

    return site_tenders


async def open_site_with_retry(browser, site):
    for attempt in range(1, MAX_OPEN_RETRIES + 1):
        try:
            print(f"{site['name']}: Opening attempt {attempt}...")
            context = await browser.new_context()
            page    = await context.new_page()
            page.set_default_timeout(site["timeout_ms"])
            await page.goto(site["url"], wait_until="domcontentloaded", timeout=site["timeout_ms"])
            return {"site": site, "context": context, "page": page, "opened": True}
        except Exception as e:
            print(f"{site['name']}: Attempt {attempt} failed: {e}")
            try: await context.close()
            except: pass
            await asyncio.sleep(2)
    return {"site": site, "context": None, "page": None, "opened": False}


async def open_all_sites(browser):
    return await asyncio.gather(*[open_site_with_retry(browser, s) for s in SITES])


async def read_portal_worker(item):
    if not item["opened"]:
        print(f"{item['site']['name']}: Skipped — portal did not open.")
        return []
    try:
        return await read_all_pages(item["page"], item["site"])
    except Exception as e:
        print(f"{item['site']['name']}: Failed while reading: {e}")
        return []


async def close_all_contexts(opened_pages):
    for item in opened_pages:
        try:
            if item["context"]:
                await item["context"].close()
        except:
            pass

# ── OUTPUT ────────────────────────────────────────────────────────────────────

def save_master_excel(tenders):
    print(f"\nTotal rows after AI filter: {len(tenders)}")
    if not tenders:
        print("No matching tenders found.")
        return

    df = pd.DataFrame(tenders)
    df = df.drop_duplicates(subset=["Tender Title / Ref / ID", "Organisation"], keep="first")
    df = df.sort_values(by=["Score", "Tender Value Number"], ascending=False)

    best_df   = df[df["Tender Type"] == "Best Tender"]
    review_df = df[df["Tender Type"] == "Review Tender"]
    backup_df = df[df["Tender Type"] == "Captured Backup"]

    df.to_excel("all_captured_1cr_tenders.xlsx", index=False)
    best_df.to_excel("best_tenders.xlsx", index=False)
    review_df.to_excel("review_tenders.xlsx", index=False)
    backup_df.to_excel("backup_1cr_tenders.xlsx", index=False)

    print("Output files saved:")
    print(f"  best_tenders.xlsx        — {len(best_df)} tenders")
    print(f"  review_tenders.xlsx      — {len(review_df)} tenders")
    print(f"  backup_1cr_tenders.xlsx  — {len(backup_df)} tenders")
    print(f"  all_captured_1cr_tenders.xlsx — {len(df)} tenders")

# ── MAIN ──────────────────────────────────────────────────────────────────────

async def main():
    all_stage1_tenders = []

    async with async_playwright() as p:
        browser      = await p.chromium.launch(headless=False, slow_mo=0)
        opened_pages = await open_all_sites(browser)

        print("\nAll portals are now open in the browser.")
        print("For each portal:")
        print("  1. Solve the captcha")
        print("  2. Click Search to load the tender list")
        print("The script will automatically detect when each portal is ready and start scraping.")
        print("You have 5 minutes per portal before it times out.")
        input("\nPress Enter to begin monitoring all portals...")

        # Stage 1: scrape + keyword filter (runs in parallel across all portals)
        results = await asyncio.gather(
            *[read_portal_worker(item) for item in opened_pages],
            return_exceptions=True
        )

        for result in results:
            if isinstance(result, Exception):
                print(f"Portal task failed: {result}")
            else:
                all_stage1_tenders.extend(result)

        await close_all_contexts(opened_pages)

    print(f"\nStage 1 complete: {len(all_stage1_tenders)} tenders passed keyword filter.")

    # Stage 2: Claude AI classifier (requires API credits — disabled for now)
    # To enable: buy credits at console.anthropic.com, set ANTHROPIC_API_KEY, then uncomment below
    # all_final_tenders = await ai_filter_tenders(all_stage1_tenders)
    all_final_tenders = all_stage1_tenders

    # Re-score and reclassify
    for t in all_final_tenders:
        score = calculate_score(t["Tender Title / Ref / ID"], t["Organisation"],
                                parse_value(t["Tender Value"]))
        t["Score"]       = score
        t["Tender Type"] = get_tender_type(score)

    # Save per-portal files (AI-filtered)
    by_portal = {}
    for t in all_final_tenders:
        by_portal.setdefault(t["Portal"], []).append(t)

    for site in SITES:
        portal_tenders = by_portal.get(site["name"], [])
        save_portal_excel(site["file"], portal_tenders)

    save_master_excel(all_final_tenders)
    print(f"\nCompleted! {len(all_final_tenders)} final tenders captured.")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())