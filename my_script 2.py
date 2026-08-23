import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import re
import urllib.parse

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FILE = "New Business Datascope 2026.xlsx"
OUTPUT_FILE = "New Business Datascope 2026.xlsx"
CAREER_COL = "Company Recruitment Website"
COUNT_COL = "Number Roles"

# ==========================================
# ATS ENGINES
# ==========================================

async def count_lever(page):
    return await page.evaluate("document.querySelectorAll('.posting').length")

async def count_jobvite(page):
    return await page.evaluate("document.querySelectorAll('.jv-job-list-item').length")

async def count_ashby(page):
    await page.wait_for_timeout(5000)
    links = await page.evaluate('''() => Array.from(document.querySelectorAll("a")).map(a => a.href)''')
    ashby_pattern = r"jobs\.ashbyhq\.com/[^/]+/[a-f0-9-]{36}"
    unique_jobs = {link for link in links if re.search(ashby_pattern, link)}
    return len(unique_jobs) if len(unique_jobs) > 0 else await page.evaluate("document.querySelectorAll('.ashby-job-posting-brief, .ashby-job-posting, .ashby-job-board-job').length")

async def count_breezy(page):
    return await page.evaluate("document.querySelectorAll('li.job, .position').length")

async def count_peopleforce(page):
    return await page.evaluate("document.querySelectorAll('.job-card, .job-listing').length")

async def count_recruitee(page):
    return await page.evaluate("document.querySelectorAll('.opening, a[href*=\"/o/\"]').length")

async def count_workable(page):
    try:
        await page.wait_for_load_state("networkidle", timeout=10000)
    except:
        await page.wait_for_timeout(5000)
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await page.wait_for_timeout(3000)
    return await page.evaluate("document.querySelectorAll('a[href*=\"/view/\"], [data-test=\"job-item\"]').length")

async def count_join(page, url):
    try:
        await page.wait_for_load_state("networkidle", timeout=10000)
    except:
        await page.wait_for_timeout(5000)
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await page.wait_for_timeout(3000)
    return await page.evaluate('''() => {
        const links = Array.from(document.querySelectorAll('a[href*="/companies/"]'));
        const jobLinks = links.filter(a => /\\/companies\\/[^/]+\\/\\d+/.test(a.href));
        const uniqueJobs = new Set(jobLinks.map(a => a.href.split('?')[0]));
        return uniqueJobs.size > 0 ? uniqueJobs.size : document.querySelectorAll('[data-testid="job-card"]').length;
    }''')

async def count_comeet(page, url):
    try:
        await page.wait_for_load_state("networkidle", timeout=10000)
    except:
        await page.wait_for_timeout(4000)
    return await page.evaluate('''() => {
        const elements = document.querySelectorAll('.comeet-position, .positionItem, [data-comeet-position], .comeet-g-position');
        if (elements.length > 0) return elements.length;
        const links = Array.from(document.querySelectorAll('a[href*="/position/"], a[href*="/jobs/"]'));
        return new Set(links.map(a => a.href.split('?')[0])).size;
    }''')

async def count_personio(page, url):
    try:
        await page.wait_for_load_state("networkidle", timeout=10000)
    except:
        await page.wait_for_timeout(4000)
    return await page.evaluate('''() => {
        const boxes = document.querySelectorAll('.job-box, .recruiting-listing-item, [id^="job-position-"], .job-position-title, [data-testid="job-list-item"]');
        if (boxes.length > 0) return boxes.length;
        const links = Array.from(document.querySelectorAll('a[href*="/job/"], a[href*="/recruiting/positions/"]'));
        return new Set(links.map(a => a.href.split('?')[0])).size;
    }''')

async def count_greenhouse(page):
    try:
        await page.wait_for_load_state("networkidle", timeout=12000)
    except:
        await page.wait_for_timeout(4000)
        
    all_links = []
    try:
        all_links.extend(await page.evaluate('''() => Array.from(document.querySelectorAll("a")).map(a => a.href)'''))
    except:
        pass
        
    # Deep extract links across any potential iframe container layouts
    for frame in page.frames:
        try:
            all_links.extend(await frame.evaluate('''() => Array.from(document.querySelectorAll("a")).map(a => a.href)'''))
        except:
            continue

    unique_jobs = set()
    for link in all_links:
        if not link: continue
        l_lower = link.lower()
        
        # Match parameter-based embedded IDs
        if "gh_jid=" in l_lower:
            match = re.search(r"gh_jid=(\d+)", l_lower)
            if match: unique_jobs.add(f"id_{match.group(1)}")
            
        elif "job_detail" in l_lower and "id=" in l_lower:
            match = re.search(r"id=(\d+)", l_lower)
            if match: unique_jobs.add(f"id_{match.group(1)}")
            
        # Match standard directory-based path IDs
        elif "/jobs/" in l_lower:
            match = re.search(r"/jobs/(\d+)", l_lower)
            if match: 
                unique_jobs.add(f"id_{match.group(1)}")
            else:
                clean = l_lower.split('?')[0].rstrip('/')
                if clean.split('/')[-1].isdigit():
                    unique_jobs.add(clean)

    if len(unique_jobs) > 0:
        return len(unique_jobs)

    # UI Element Layout fallbacks if tracking tags are masked
    return await page.evaluate('''() => {
        const openings = document.querySelectorAll('.opening');
        if (openings.length > 0) return openings.length;
        return document.querySelectorAll('[data-testid="job-row"], .job-post, .job-row, [id^="job-row-"]').length;
    }''')

async def count_generic(page):
    try:
        await page.wait_for_load_state("networkidle", timeout=15000)
    except:
        await page.wait_for_timeout(5000)
    
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await page.wait_for_timeout(3000)
    
    all_links = []
    all_links.extend(await page.evaluate('''() => Array.from(document.querySelectorAll("a")).map(a => a.href)'''))
    
    for _ in range(2):
        for frame in page.frames:
            try:
                all_links.extend(await frame.evaluate('''() => Array.from(document.querySelectorAll("a")).map(a => a.href)'''))
            except: continue
        await asyncio.sleep(1)
        
    unique_links = set()
    if "notion.site" in page.url.lower():
        return 1
    
    patterns = [
        "/job/", "/jobs/", "/vacancy/", "/postings/", 
        "greenhouse.io/", "/j/", "lever.co/", "ashbyhq.com/", 
        "apply/", "gh_jid=", "notion.site/", "/o/", "/career/", 
        "/careers/", "/position/", "/positions/", "/recruiting/"
    ]
    
    for link in all_links:
        if not link: continue
        if any(p in link.lower() for p in patterns):
            clean = link.split('?')[0].rstrip('/').lower()
            unique_links.add(clean)
            
    if len(unique_links) > 0:
        return len(unique_links)
        
    count = await page.evaluate('''() => {
        const elements = Array.from(document.querySelectorAll('a, button, div, span'));
        const applyElements = elements.filter(el => {
            const text = el.innerText.trim().toLowerCase();
            return text === 'apply now' || text === 'apply' || text === 'view job' || text === 'view opening';
        });
        return new Set(applyElements.map(el => el.closest('a') || el)).size;
    }''')
    
    return count

# ==========================================
# SMART DISPATCHER
# ==========================================
async def get_count(page, url):
    u = url.lower()
    
    # 1. Direct Deep-Link Protections (Safe Singular Bypasses)
    if "gh_jid=" in u: 
        return 1
    if "join.com" in u and re.search(r"/companies/[^/]+/\d+", u): 
        return 1
    if "comeet" in u and (re.search(r"/position/[^/]+$", u) or re.search(r"/jobs/[^/]+/[^/]+/[^/]+", u)): 
        return 1
    if "personio" in u and re.search(r"/job/\d+", u): 
        return 1
    
    # 2. Routing Engines
    count = 0
    if "greenhouse.io" in u: count = await count_greenhouse(page)
    elif "lever.co" in u: count = await count_lever(page)
    elif "jobvite.com" in u: count = await count_jobvite(page)
    elif "ashbyhq.com" in u: count = await count_ashby(page)
    elif "breezy.hr" in u: count = await count_breezy(page)
    elif "peopleforce.io" in u: count = await count_peopleforce(page)
    elif "recruitee.com" in u or "jobs." in u: count = await count_recruitee(page)
    elif "workable.com" in u: count = await count_workable(page)
    elif "join.com" in u: count = await count_join(page, url)
    elif "comeet" in u: count = await count_comeet(page, url)
    elif "personio" in u: count = await count_personio(page, url)
    
    if count == 0:
        count = await count_generic(page)
        
    return count

# ==========================================
# RUNNER
# ==========================================
async def worker(idx, row, browser, df, semaphore, write_lock, progress_counter):
    url = str(row.get(CAREER_COL, "")).strip()
    if not url or url.lower() == "nan" or url == "": return

    async with semaphore:
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080}
        )
        page = await context.new_page()
        await page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            count = await get_count(page, url)
        except Exception as e:
            print(f"Error on {url}: {e}")
            count = 0
        finally:
            await context.close()

        async with write_lock:
            df.loc[idx, COUNT_COL] = count
            progress_counter["count"] += 1
            print(f"[{progress_counter['count']}] {url} -> {count}")
            if progress_counter["count"] % 5 == 0: df.to_excel(OUTPUT_FILE, index=False)

async def main():
    df = pd.read_excel(INPUT_FILE)
    semaphore = asyncio.Semaphore(2) 
    write_lock = asyncio.Lock()
    progress_counter = {"count": 0}
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        tasks = [worker(idx, row, browser, df, semaphore, write_lock, progress_counter) for idx, row in df.iterrows()]
        await asyncio.gather(*tasks)
        await browser.close()
    
    df.to_excel(OUTPUT_FILE, index=False)
    print("Done!")

if __name__ == "__main__":
    asyncio.run(main())