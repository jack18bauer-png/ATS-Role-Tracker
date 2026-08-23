"""
scan_jobs.py

Counts open roles on company career sites and writes the number into the
BD spreadsheet.

Two-layer approach:

  Layer 1 (fast, deterministic): aiohttp against the ATS public JSON/XML
  endpoints. No browser, no DOM guessing, no rendering race conditions.
  If the API answers, that number is the truth.

  Layer 2 (fallback): Playwright, for platforms with no usable public API
  and for bespoke career pages.

Before either layer runs, a Tier 0 empty-state check looks for "no open
positions" style copy so that a genuinely empty board records 0 instead of
being handed to the generic link-pattern counter, which will happily count
nav links and return nonsense.

Requirements:
    pip install pandas openpyxl aiohttp playwright
    python -m playwright install chromium

Run:
    python scan_jobs.py
"""

import asyncio
import re
import sys
import xml.etree.ElementTree as ET
from urllib.parse import urlparse

import aiohttp
import pandas as pd
from playwright.async_api import async_playwright

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FILE = "New Business Datascope 2026.xlsx"
OUTPUT_FILE = "New Business Datascope 2026.xlsx"
CAREER_COL = "Company Recruitment Website"
COUNT_COL = "Number Roles"
METHOD_COL = "Count Method"          # audit trail: which layer produced the number

CONCURRENCY = 4                       # API calls are cheap, browser pages are not
BROWSER_CONCURRENCY = 2
SAVE_EVERY = 5
API_TIMEOUT = 20
PAGE_TIMEOUT = 60000

# Any generic-fallback count above this is almost certainly counting
# navigation or footer links rather than jobs. Flagged for manual review
# rather than silently written.
MAX_PLAUSIBLE_COUNT = 50

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

# ==========================================
# TIER 0: EMPTY STATE DETECTION
# ==========================================
EMPTY_STATE_PATTERNS = [
    r"no\s+(current\s+|currently\s+|open\s+|active\s+)?(job\s+)?(openings|opportunities|positions|vacancies|roles|jobs)",
    r"there\s+are\s+(currently\s+)?no\s+",
    r"we\s+(don'?t|do\s+not)\s+have\s+any\s+(open\s+)?(positions|roles|vacancies)",
    r"0\s+(open\s+)?(positions|roles|jobs|vacancies)",
    r"no\s+results\s+(were\s+)?found",
    r"no\s+matching\s+jobs",
    r"check\s+back\s+(soon|later)",
    r"position(s)?\s+are\s+not\s+available",
]
EMPTY_STATE_RE = re.compile("|".join(EMPTY_STATE_PATTERNS), re.IGNORECASE)


def looks_empty(text: str) -> bool:
    """True if the page copy says there are no roles."""
    if not text:
        return False
    # Only test a trimmed body. Long pages with blog content produce false hits.
    return bool(EMPTY_STATE_RE.search(text[:6000]))


# Sites occasionally sit behind a login wall (Cloudflare Access, etc). In
# that case the page loads (200 OK) but there's no job content to see, and
# writing 0 would misrepresent "couldn't check" as "confirmed zero roles".
AUTH_WALL_PATTERNS = [
    r"log\s*in\s+to\s+[\w.\-]+",
    r"sign\s*in\s+to\s+continue",
    r"authentication\s+required",
    r"you\s+need\s+to\s+(log|sign)\s*in\s+to",
    r"send\s+login\s+code",
]
AUTH_WALL_RE = re.compile("|".join(AUTH_WALL_PATTERNS), re.IGNORECASE)


def looks_authwalled(page_url: str, text: str) -> bool:
    """True if the page is a login gate rather than the real career page."""
    if "cloudflareaccess.com" in (page_url or "").lower():
        return True
    if not text:
        return False
    return bool(AUTH_WALL_RE.search(text[:3000]))


# ==========================================
# PER-SITE RULES
# ==========================================
# For career pages where the generic counter cannot be trusted. Keys are
# matched as substrings against the lowercased URL.
#
#   allow: link must match at least one of these regexes to count
#   deny:  link matching any of these is discarded
#   cap:   hard ceiling, useful where a site paginates oddly
SITE_RULES = {
    "g5.com": {
        # G5 Games: strict allowlist. The careers page links out to a large
        # number of non-job pages that all sit under /careers/.
        "allow": [r"/careers/vacancy/", r"/vacancies?/\d+"],
        "deny": [r"/careers/?$", r"/careers/(life|benefits|culture|about)"],
    },
    "kalypsomedia.com": {
        "allow": [r"[?&]id=\d+"],
    },
}


def rules_for(url: str):
    u = url.lower()
    for key, rules in SITE_RULES.items():
        if key in u:
            return rules
    return None


def apply_rules(links, rules):
    if not rules:
        return links
    out = []
    for link in links:
        low = link.lower()
        if rules.get("deny") and any(re.search(p, low) for p in rules["deny"]):
            continue
        if rules.get("allow") and not any(re.search(p, low) for p in rules["allow"]):
            continue
        out.append(link)
    if rules.get("cap"):
        out = out[: rules["cap"]]
    return out


# ==========================================
# LAYER 1: ATS API HANDLERS
# ==========================================
# Each returns an int, or None if it could not answer (bad token, 404,
# network error). None means "fall through to the browser", 0 means
# "confirmed zero open roles".

def _host(url):
    try:
        return (urlparse(url).hostname or "").lower()
    except Exception:
        return ""


async def _get_json(session, url):
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=API_TIMEOUT)) as r:
            if r.status != 200:
                return None
            return await r.json(content_type=None)
    except Exception:
        return None


async def _get_text(session, url):
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=API_TIMEOUT)) as r:
            if r.status != 200:
                return None
            return await r.text()
    except Exception:
        return None


async def api_greenhouse(session, url):
    m = re.search(r"greenhouse\.io/(?:embed/job_board\?for=)?([a-z0-9_\-]+)", url, re.I)
    if not m:
        m = re.search(r"[?&]for=([a-z0-9_\-]+)", url, re.I)
    if not m:
        return None
    token = m.group(1)
    if token in {"embed", "boards", "job-boards"}:
        return None
    for base in ("https://boards-api.greenhouse.io/v1/boards",
                 "https://boards-api.eu.greenhouse.io/v1/boards"):
        data = await _get_json(session, f"{base}/{token}/jobs")
        if isinstance(data, dict) and "jobs" in data:
            return len(data["jobs"])
    return None


async def api_lever(session, url):
    m = re.search(r"lever\.co/([a-z0-9_\-\.]+)", url, re.I)
    if not m:
        return None
    data = await _get_json(session, f"https://api.lever.co/v0/postings/{m.group(1)}?mode=json")
    if isinstance(data, list):
        return len(data)
    return None


async def api_ashby(session, url):
    m = re.search(r"ashbyhq\.com/([a-z0-9_\-\.%20\+]+)", url, re.I)
    if not m:
        return None
    token = m.group(1).split("/")[0]
    data = await _get_json(
        session, f"https://api.ashbyhq.com/posting-api/job-board/{token}"
    )
    # The payload key is "jobs". An earlier version of this script looked for
    # "postings" and silently returned 0 for every Ashby board.
    if isinstance(data, dict) and isinstance(data.get("jobs"), list):
        return len(data["jobs"])
    return None


async def api_workable(session, url):
    m = re.search(r"(?:apply\.workable\.com/(?:j/)?|//)([a-z0-9_\-]+)\.?workable\.com", url, re.I)
    m2 = re.search(r"apply\.workable\.com/([a-z0-9_\-]+)", url, re.I)
    token = None
    if m2:
        token = m2.group(1)
    elif m:
        token = m.group(1)
    if not token or token in {"apply", "www"}:
        return None
    data = await _get_json(
        session,
        f"https://apply.workable.com/api/v1/widget/accounts/{token}?details=false",
    )
    if isinstance(data, dict) and isinstance(data.get("jobs"), list):
        return len(data["jobs"])
    return None


async def api_recruitee(session, url):
    m = re.search(r"([a-z0-9_\-]+)\.recruitee\.com", url, re.I)
    if not m:
        return None
    data = await _get_json(session, f"https://{m.group(1)}.recruitee.com/api/offers/")
    if isinstance(data, dict) and isinstance(data.get("offers"), list):
        return len(data["offers"])
    return None


async def api_personio(session, url):
    m = re.search(r"([a-z0-9_\-]+)\.jobs\.personio\.(?:de|com)", url, re.I)
    if not m:
        return None
    for tld in ("de", "com"):
        xml = await _get_text(session, f"https://{m.group(1)}.jobs.personio.{tld}/xml")
        if not xml:
            continue
        try:
            root = ET.fromstring(xml)
            return len(root.findall(".//position"))
        except ET.ParseError:
            continue
    # Some tenants have been migrated to Personio's newer Next.js career
    # site, which drops the classic /xml feed (404s). That site still
    # server-renders each posting as a plain /job/<id> link in the HTML,
    # so count those directly instead of falling through to the browser.
    for tld in ("de", "com"):
        html = await _get_text(session, f"https://{m.group(1)}.jobs.personio.{tld}/")
        if not html:
            continue
        ids = set(re.findall(r'/job/(\d+)', html))
        if ids:
            return len(ids)
    return None


async def api_teamtailor(session, url):
    m = re.search(r"([a-z0-9_\-]+)\.teamtailor\.com", url, re.I)
    if not m:
        return None
    rss = await _get_text(session, f"https://{m.group(1)}.teamtailor.com/jobs.rss")
    if not rss:
        return None
    try:
        root = ET.fromstring(rss)
        return len(root.findall(".//item"))
    except ET.ParseError:
        return None


async def api_smartrecruiters(session, url):
    m = re.search(r"smartrecruiters\.com/([a-zA-Z0-9_\-]+)", url)
    if not m:
        return None
    token = m.group(1)
    data = await _get_json(
        session, f"https://api.smartrecruiters.com/v1/companies/{token}/postings?limit=100"
    )
    if isinstance(data, dict):
        if isinstance(data.get("totalFound"), int):
            return data["totalFound"]
        if isinstance(data.get("content"), list):
            return len(data["content"])
    return None


async def api_bamboohr(session, url):
    m = re.search(r"([a-z0-9_\-]+)\.bamboohr\.com", url, re.I)
    if not m:
        return None
    data = await _get_json(session, f"https://{m.group(1)}.bamboohr.com/careers/list")
    if isinstance(data, dict) and isinstance(data.get("result"), list):
        return len(data["result"])
    return None


async def api_breezy(session, url):
    m = re.search(r"([a-z0-9_\-]+)\.breezy\.hr", url, re.I)
    if not m:
        return None
    data = await _get_json(session, f"https://{m.group(1)}.breezy.hr/json")
    if isinstance(data, list):
        return len(data)
    return None


async def api_jobvite(session, url):
    m = re.search(r"jobvite\.com/(?:careers/)?([a-z0-9_\-]+)", url, re.I)
    if not m:
        return None
    data = await _get_json(
        session, f"https://api.jobvite.com/api/v2/job?companyId={m.group(1)}"
    )
    if isinstance(data, dict) and isinstance(data.get("requisitions"), list):
        return len(data["requisitions"])
    return None


API_HANDLERS = [
    ("greenhouse.io", api_greenhouse),
    ("lever.co", api_lever),
    ("ashbyhq.com", api_ashby),
    ("workable.com", api_workable),
    ("recruitee.com", api_recruitee),
    ("personio.", api_personio),
    ("teamtailor.com", api_teamtailor),
    ("smartrecruiters.com", api_smartrecruiters),
    ("bamboohr.com", api_bamboohr),
    ("breezy.hr", api_breezy),
    ("jobvite.com", api_jobvite),
]


async def try_api(session, url):
    """Returns (count, method) or (None, None)."""
    u = url.lower()
    for key, handler in API_HANDLERS:
        if key in u:
            try:
                count = await handler(session, url)
            except Exception:
                count = None
            if count is not None:
                return count, f"api:{key.rstrip('.')}"
            return None, None
    return None, None


# ==========================================
# LAYER 2: BROWSER HANDLERS
# ==========================================

async def _settle(page, timeout=10000, scroll=False):
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        await page.wait_for_timeout(4000)
    if scroll:
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await page.wait_for_timeout(2500)


async def _all_links(page):
    links = []
    try:
        links.extend(await page.evaluate(
            '() => Array.from(document.querySelectorAll("a")).map(a => a.href)'
        ))
    except Exception:
        pass
    for frame in page.frames:
        try:
            links.extend(await frame.evaluate(
                '() => Array.from(document.querySelectorAll("a")).map(a => a.href)'
            ))
        except Exception:
            continue
    return [l for l in links if l]


async def count_peopleforce(page):
    await _settle(page)
    return await page.evaluate(r'''() => {
        const cards = document.querySelectorAll('.job-card, .job-listing-item');
        if (cards.length > 0) return cards.length;
        const links = Array.from(document.querySelectorAll('a[href*="/careers/positions/"]'));
        return new Set(links.map(a => a.href.split('?')[0].replace(/\/$/, '').toLowerCase())).size;
    }''')


async def count_join(page):
    await _settle(page, scroll=True)
    return await page.evaluate(r'''() => {
        const links = Array.from(document.querySelectorAll('a[href*="/companies/"]'));
        const jobLinks = links.filter(a => /\/companies\/[^/]+\/\d+/.test(a.href));
        const unique = new Set(jobLinks.map(a => a.href.split('?')[0]));
        return unique.size > 0 ? unique.size : document.querySelectorAll('[data-testid="job-card"]').length;
    }''')


async def count_comeet(page):
    await _settle(page)
    return await page.evaluate(r'''() => {
        const els = document.querySelectorAll('.comeet-position, .positionItem, [data-comeet-position], .comeet-g-position');
        if (els.length > 0) return els.length;
        const links = Array.from(document.querySelectorAll('a[href*="/position/"], a[href*="/jobs/"]'));
        return new Set(links.map(a => a.href.split('?')[0])).size;
    }''')


async def count_generic(page, url):
    await _settle(page, timeout=15000, scroll=True)

    if "notion.site" in page.url.lower():
        return 1

    links = await _all_links(page)

    rules = rules_for(url)
    if rules:
        links = apply_rules(links, rules)
        return len({l.split("#")[0].split("?")[0].rstrip("/").lower() for l in links}) if not rules.get("allow") \
            else len({l.split("#")[0].rstrip("/").lower() for l in links})

    # "/career/" and "/careers/" are deliberately excluded: they match the
    # careers page's own nav link and in-page anchors (e.g. #open-positions)
    # back to itself, which inflates the count on pages with no real job
    # links to find. Individual postings are caught by the more specific
    # patterns below.
    patterns = [
        "/job/", "/jobs/", "/vacancy/", "/vacancies/", "/postings/",
        "greenhouse.io/", "/j/", "lever.co/", "ashbyhq.com/",
        "apply/", "gh_jid=", "/o/",
        "/position/", "/positions/", "/recruiting/", "/opening/",
    ]
    unique = set()
    for link in links:
        low = link.lower()
        if any(p in low for p in patterns):
            # Strip the fragment too, not just the query string, so
            # "/careers/#open-positions" and "/careers/#page" (two anchors
            # into the same page) don't get counted as separate jobs.
            unique.add(low.split("#")[0].split("?")[0].rstrip("/"))

    if unique:
        return len(unique)

    return await page.evaluate(r'''() => {
        const els = Array.from(document.querySelectorAll('a, button, div, span'));
        const apply = els.filter(el => {
            const t = (el.innerText || '').trim().toLowerCase();
            // 'apply' / 'apply now' are deliberately excluded: they're also
            // the standard text for a standing "send an open application"
            // CTA that career pages show even when they have zero current
            // openings, so counting them produces phantom roles. 'view job'
            // / 'view opening' only ever appear attached to an actual
            // posting, so they stay.
            return t === 'view job' || t === 'view opening';
        });
        // innerText bubbles up through wrapper elements, so a single button
        // like <a><div><span>View job</span></div></a> matches at every
        // level. Keep only the innermost match per chain so one visual
        // button counts once instead of three or four times.
        const leafOnly = apply.filter(el => !apply.some(other => other !== el && el.contains(other)));
        return new Set(leafOnly.map(el => el.closest('a') || el)).size;
    }''')


# Single job pages. Counting links on these returns garbage, so short-circuit.
DEEP_LINK_PATTERNS = [
    r"gh_jid=",
    r"join\.com/companies/[^/]+/\d+",
    r"comeet\.com/jobs/[^/]+/[^/]+/[^/]+",
    r"comeet\.com/[^/]*/position/",
    r"personio\.[a-z]+/job/\d+",
    r"kalypsomedia\.com/.*[?&]id=\d+",
    r"lever\.co/[^/]+/[a-f0-9\-]{36}",
    r"ashbyhq\.com/[^/]+/[a-f0-9\-]{36}",
]


def is_deep_link(url):
    low = url.lower()
    return any(re.search(p, low) for p in DEEP_LINK_PATTERNS)


async def browser_count(page, url):
    """Returns (count, method)."""
    u = url.lower()

    # Tier 0 before anything else.
    try:
        body = await page.evaluate("() => document.body ? document.body.innerText : ''")
    except Exception:
        body = ""
    if looks_authwalled(page.url, body):
        return 0, "blocked:authwall REVIEW"
    if looks_empty(body):
        return 0, "tier0:empty"

    if "peopleforce.io" in u:
        return await count_peopleforce(page), "dom:peopleforce"
    if "join.com" in u:
        return await count_join(page), "dom:join"
    if "comeet" in u:
        return await count_comeet(page), "dom:comeet"

    count = await count_generic(page, url)
    method = "dom:generic"

    if count > MAX_PLAUSIBLE_COUNT:
        method = f"dom:generic REVIEW({count})"

    return count, method


# ==========================================
# RUNNER
# ==========================================

async def process_row(idx, url, session, browser, api_sem, browser_sem,
                      df, write_lock, progress):
    count, method = None, None

    if is_deep_link(url):
        count, method = 1, "deeplink:single"

    if count is None:
        async with api_sem:
            count, method = await try_api(session, url)

    if count is None:
        async with browser_sem:
            context = await browser.new_context(
                user_agent=USER_AGENT,
                viewport={"width": 1920, "height": 1080},
            )
            page = await context.new_page()
            await page.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)
                count, method = await browser_count(page, url)
            except Exception as e:
                count, method = 0, f"error:{type(e).__name__}"
            finally:
                await context.close()

    async with write_lock:
        df.loc[idx, COUNT_COL] = count
        df.loc[idx, METHOD_COL] = method
        progress["done"] += 1
        flag = "  <-- CHECK" if method and "REVIEW" in method else ""
        print(f"[{progress['done']}/{progress['total']}] {count:>4}  {method:<28} {url}{flag}")
        if progress["done"] % SAVE_EVERY == 0:
            try:
                df.to_excel(OUTPUT_FILE, index=False)
            except PermissionError:
                print("  ! Could not save, the workbook is open in Excel. Close it.")


async def main():
    df = pd.read_excel(INPUT_FILE)

    if CAREER_COL not in df.columns:
        print(f"Column '{CAREER_COL}' not found. Columns are: {list(df.columns)}")
        sys.exit(1)
    if METHOD_COL not in df.columns:
        df[METHOD_COL] = ""

    # Optional: python "scan_jobs.py" also accepts a row-limit CLI arg to
    # restrict the scan to the first N spreadsheet rows, e.g. `python
    # scan_jobs.py 20`. Omit it (normal usage) and every row is scanned.
    row_limit = None
    if len(sys.argv) > 1:
        try:
            row_limit = int(sys.argv[1])
        except ValueError:
            print(f"Ignoring invalid row limit arg: {sys.argv[1]!r}")

    scan_df = df.head(row_limit) if row_limit else df

    targets = []
    for idx, row in scan_df.iterrows():
        url = str(row.get(CAREER_COL, "")).strip()
        if url and url.lower() != "nan" and url.startswith("http"):
            targets.append((idx, url))

    print(f"{len(targets)} URLs to scan\n")

    api_sem = asyncio.Semaphore(CONCURRENCY)
    browser_sem = asyncio.Semaphore(BROWSER_CONCURRENCY)
    write_lock = asyncio.Lock()
    progress = {"done": 0, "total": len(targets)}

    headers = {"User-Agent": USER_AGENT, "Accept": "application/json, text/xml, */*"}
    connector = aiohttp.TCPConnector(limit=CONCURRENCY * 2, ssl=False)

    async with aiohttp.ClientSession(headers=headers, connector=connector) as session:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            tasks = [
                process_row(idx, url, session, browser, api_sem, browser_sem,
                            df, write_lock, progress)
                for idx, url in targets
            ]
            await asyncio.gather(*tasks)
            await browser.close()

    try:
        df.to_excel(OUTPUT_FILE, index=False)
    except PermissionError:
        df.to_excel("New Business Datascope 2026 (scan output).xlsx", index=False)
        print("Workbook was locked, wrote to 'New Business Datascope 2026 (scan output).xlsx' instead.")

    api_rows = df[METHOD_COL].astype(str).str.startswith("api:").sum()
    review_rows = df[METHOD_COL].astype(str).str.contains("REVIEW").sum()
    error_rows = df[METHOD_COL].astype(str).str.startswith("error:").sum()
    print(f"\nDone. {api_rows} via API, {review_rows} flagged for review, {error_rows} errors.")


if __name__ == "__main__":
    asyncio.run(main())