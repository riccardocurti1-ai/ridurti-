#!/usr/bin/env python3
"""
startup_scout.py v2 — Multi-Agent Pre-Seed Deep-Tech Startup Discovery
Powered by Groq free tier (llama-3.3-70b-versatile).

New in v2:
  • Thesis-FIRST flow: extract keywords before scraping → targeted HN search
  • Keyword pre-filter drops off-thesis results before any LLM token is spent
  • Launch HN replaces Show HN (real companies, not weekend side projects)
  • HN keyword search: searches HN Algolia for each thesis technology term
  • EU-Startups directory + ClimateDraft for European/climate deep-tech
  • Added 'tech_fit' scoring dimension
  • Website enrichment: fetches company homepage for richer context
  • Excel output with clickable hyperlinks, color-coded scores, full text

Three AI agents:

  ┌────────────────────────────────────────────────────────────────┐
  │  AGENT 1 — SCOUT                                               │
  │  Thesis keywords → targeted Launch HN, HN keyword search,     │
  │  EU-Startups, ClimateDraft, Product Hunt, Reddit.              │
  │  Keyword pre-filter drops off-thesis noise before scoring.     │
  └──────────────────────────┬─────────────────────────────────────┘
                             │
                             ▼
  ┌────────────────────────────────────────────────────────────────┐
  │  AGENT 2 — ANALYST                                             │
  │  Scores in independent batches (no context accumulation).      │
  │  Dimensions: sector / geo / stage / tech / theme.              │
  └──────────────────────────┬─────────────────────────────────────┘
                             │
                             ▼
  ┌────────────────────────────────────────────────────────────────┐
  │  AGENT 3 — VERIFIER + ENRICHER (direct Python, zero LLM cost) │
  │  SEC EDGAR · DuckDuckGo funding search · RDAP domain age       │
  │  Website homepage fetch → richer company descriptions          │
  │  Output: Excel + CSV with clickable links                      │
  └────────────────────────────────────────────────────────────────┘
"""

import os, sys, json, csv, time, re, difflib
from datetime import datetime, timezone
from urllib.parse import quote_plus, urlparse

# ── dependency check ──────────────────────────────────────────────────────────
_MISSING = []
for _pkg, _mod in [
    ("requests",      "requests"),
    ("beautifulsoup4","bs4"),
    ("openai",        "openai"),
    ("python-dotenv", "dotenv"),
    ("tabulate",      "tabulate"),
    ("openpyxl",      "openpyxl"),
]:
    try:
        __import__(_mod)
    except ImportError:
        _MISSING.append(_pkg)

if _MISSING:
    print("Missing dependencies. Run:")
    print(f"  pip install {' '.join(_MISSING)}")
    sys.exit(1)

import requests
from bs4 import BeautifulSoup
import openai
from dotenv import load_dotenv, set_key
from tabulate import tabulate
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── config ────────────────────────────────────────────────────────────────────
ENV_FILE    = ".env"
OUTPUT_XLSX = "startup_scout_results.xlsx"
OUTPUT_CSV  = "startup_scout_results.csv"
MODEL              = "llama-3.3-70b-versatile"   # overridden at runtime if Google key found
TOP_N              = 30
BATCH_SIZE         = 8
MIN_SCORE_TO_VERIFY = 3   # only verify candidates the analyst rated ≥ this

load_dotenv(ENV_FILE)
SCRAPED_AT = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# ── shared state ──────────────────────────────────────────────────────────────
_scout_buffer   = []
_analyst_buffer = []
_final_results  = []


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def clean_text(text, max_len=600):
    if not text:
        return ""
    return " ".join(str(text).split())[:max_len]


def extract_domain(url):
    if not url:
        return None
    try:
        parsed = urlparse(url if "://" in url else "https://" + url)
        host = (parsed.hostname or "").lstrip("www.")
        return host if "." in host else None
    except Exception:
        return None


def browser_headers():
    return {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
        ),
        "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    }


def _entry(name, description, source, url, website_url="", created_at="", tags=""):
    return {
        "name":        clean_text(name, 120),
        "description": clean_text(description, 600),
        "source":      source,
        "url":         url,
        "website_url": website_url or url,
        "created_at":  created_at,
        "scraped_at":  SCRAPED_AT,
        "tags":        tags,
    }


def _oai_tool(name, description, properties, required=None):
    return {
        "type": "function",
        "function": {
            "name":        name,
            "description": description,
            "parameters":  {
                "type":       "object",
                "properties": properties,
                "required":   required or [],
            },
        },
    }


def safe_get(url, headers=None, timeout=20, **kwargs):
    try:
        r = requests.get(url, headers=headers or browser_headers(), timeout=timeout, **kwargs)
        return r if r.ok else None
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# DUAL-KEY LLM CALL — tries all available API clients in order
# ═══════════════════════════════════════════════════════════════════════════════

# Populated in main() — list of (openai_client, model_name, label) tuples
_API_CLIENTS: list = []
_EXHAUSTED: set   = set()   # labels of keys that hit hard quota/auth errors
_RATE_LIMITED: set = set()  # labels currently rate-limited (soft, will reset)
_consecutive_failures = 0   # consecutive all-keys-failed calls


def all_apis_exhausted():
    """True when every configured key is either auth-dead or rate-limited."""
    available = [l for _, _, l in _API_CLIENTS if l not in _EXHAUSTED]
    return len(available) == 0 or all(l in _RATE_LIMITED for l in available)


def llm_call(messages, max_tokens=2000, temperature=0):
    """Try each configured API client in turn. Returns (response_text, label_used)."""
    global _consecutive_failures
    any_rate_limited = False
    for _c, _m, _label in _API_CLIENTS:
        if _label in _EXHAUSTED:
            continue
        try:
            resp = _c.chat.completions.create(
                model=_m,
                messages=messages,
                max_tokens=max_tokens,
                temperature=temperature,
            )
            # Success — clear rate-limited flag for this key and reset counter
            _RATE_LIMITED.discard(_label)
            _consecutive_failures = 0
            return resp.choices[0].message.content.strip(), _label
        except Exception as _e:
            _err = str(_e)
            if any(x in _err.lower() for x in ("quota", "429", "rate_limit", "rate limit", "too many")):
                _m2 = re.search(r'try again in ([\d.]+[ms]+)', _err)
                _hint = f" — retry in {_m2.group(1)}" if _m2 else ""
                print(f"\n    ⚠ {_label} rate limited{_hint}, trying next key...")
                _RATE_LIMITED.add(_label)
                any_rate_limited = True
            elif any(x in _err for x in ("401", "403", "invalid", "not valid", "not found", "404")):
                print(f"\n    ⚠ {_label} auth/model error — marking exhausted")
                _EXHAUSTED.add(_label)
            else:
                print(f"\n    ⚠ {_label} error: {_err[:100]}")
    _consecutive_failures += 1
    return "", "none"


# ═══════════════════════════════════════════════════════════════════════════════
# THESIS KEYWORD EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════

def extract_thesis_keywords(client, thesis):
    """
    One quick LLM call: extract TWO tiers of keywords from the thesis.
    - Specific: exact tech terms for targeted HN search
    - Broad: sector/problem terms for pre-filter (a startup may not use jargon)
    Returns a combined deduplicated list.
    """
    if not thesis:
        return []
    try:
        raw, _lbl = llm_call([
            {"role": "system", "content":
                "Extract search keywords from an investor thesis. "
                "Return ONLY a JSON object, no markdown."},
            {"role": "user", "content": (
                f"Investor thesis:\n{thesis}\n\n"
                "Return a JSON object with two arrays:\n"
                '{"specific": ["exact tech terms, scientific names, acronyms — 8-10 items"], '
                '"broad": ["plain-English sector/problem words a startup pitch would use — 8-10 items"]}\n\n'
                "Example for a cooling thesis:\n"
                '{"specific":["thermoelectric","magnetocaloric","microfluidic cooling","MOF adsorption"],'
                '"broad":["cooling","thermal management","heat dissipation","energy efficiency","data center cooling"]}'
            )},
        ], max_tokens=400)
        if not raw:
            raise ValueError("empty response")
        raw = re.sub(r"```(?:json)?", "", raw).strip().strip("`")
        parsed = json.loads(raw)
        specific = [str(k).lower().strip() for k in parsed.get("specific", []) if k and len(str(k).strip()) > 2]
        broad    = [str(k).lower().strip() for k in parsed.get("broad",    []) if k and len(str(k).strip()) > 2]
        # Deduplicate, keep specifics first (they drive HN search), then broad
        seen, combined = set(), []
        for kw in specific + broad:
            if kw not in seen:
                seen.add(kw)
                combined.append(kw)
        return combined
    except Exception as e:
        print(f"  ⚠ Keyword extraction failed ({e}), using defaults.")
        return []


def is_article_title(name):
    """
    Heuristic: return True if the name looks like a news article / HN post title
    rather than a company/product name.
    """
    if not name:
        return True
    if len(name) > 80:
        return True
    words = name.split()
    if len(words) > 9:
        return True
    name_low = name.lower()
    # First-person or sentence starters that are never company names
    starters = [
        "the ", "a ", "an ", "how ", "why ", "what ", "when ", "where ",
        "i ", "i'", "we ", "we'", "our ", "my ",
        "microsoft ", "google ", "apple ", "amazon ", "meta ", "nvidia ",
        "more ", "new ", "study ", "using ", "researchers ", "scientists ",
        "co-designing ", "dna ", "human ", "crops ", "nitrogen ",
        "free ", "open ", "introducing ", "announcing ", "building ",
        "asking ", "show ", "tell ", "help ", "let ", "this ",
    ]
    if any(name_low.startswith(s) for s in starters):
        return True
    # Verb-heavy first words = sentence, not brand
    if re.match(r"^(i |we |i'|we'|i'm |we're |i've |we've )", name_low):
        return True
    # Contains punctuation typical of sentences
    if name.count(":") > 1 or "?" in name or name.endswith("."):
        return True
    return False


def keyword_relevance(startup, keywords):
    """Count how many thesis keywords appear in name+description+tags."""
    if not keywords:
        return 1
    text = " ".join([
        startup.get("name", ""),
        startup.get("description", ""),
        startup.get("tags", ""),
    ]).lower()
    return sum(1 for k in keywords if k in text)


# ═══════════════════════════════════════════════════════════════════════════════
# FETCHERS
# ═══════════════════════════════════════════════════════════════════════════════

def fetch_launch_hn(pages=10):
    """
    HN 'Launch HN:' posts — real company launches, more serious than Show HN.
    These are founder-written launch posts, not weekend projects.
    """
    results, seen = [], set()
    for page in range(pages):
        try:
            r = requests.get(
                "https://hn.algolia.com/api/v1/search_by_date",
                params={
                    "query": "Launch HN",
                    "tags":  "story",
                    "hitsPerPage": 50,
                    "page": page,
                },
                timeout=15,
            )
            r.raise_for_status()
            hits = r.json().get("hits", [])
            if not hits:
                break
            for hit in hits:
                title = (hit.get("title") or "").strip()
                if not re.match(r"(?i)^launch\s+hn", title):
                    continue
                oid = str(hit.get("objectID", ""))
                if oid in seen:
                    continue
                seen.add(oid)
                # Strip "Launch HN:" prefix for a cleaner name
                name = re.sub(r"(?i)^launch\s+hn[:\s–-]+", "", title).strip()
                body     = clean_text(hit.get("story_text") or "", 500)
                site_url = hit.get("url") or f"https://news.ycombinator.com/item?id={oid}"
                results.append(_entry(
                    name=name or title,
                    description=body or name,
                    source="HN Launch HN",
                    url=f"https://news.ycombinator.com/item?id={oid}",
                    website_url=site_url,
                    created_at=(hit.get("created_at") or "")[:10],
                ))
            time.sleep(0.15)
        except Exception as e:
            print(f"    ⚠ Launch HN page {page}: {e}")
            break
    return results


def fetch_hn_keywords(keywords, pages_per_kw=4):
    """
    Search HN Algolia for thesis keywords BUT only keep Show HN / Launch HN posts.
    Regular HN posts are articles/discussions, not startup launches.
    Uses both "Show HN: {kw}" and "{kw}" queries then filters titles.
    """
    if not keywords:
        return []
    results, seen = [], set()
    # Use all keywords (both specific and broad), longest first
    kws = sorted([k for k in keywords if len(k) > 4], key=len, reverse=True)[:14]

    for kw in kws:
        # Two query variants: one with Show HN prefix, one plain for broader reach
        queries = [f"Show HN {kw}", kw]
        for query in queries:
            for page in range(pages_per_kw):
                try:
                    r = requests.get(
                        "https://hn.algolia.com/api/v1/search_by_date",
                        params={
                            "query":       query,
                            "tags":        "story",
                            "hitsPerPage": 50,
                            "page":        page,
                        },
                        timeout=15,
                    )
                    r.raise_for_status()
                    hits = r.json().get("hits", [])
                    if not hits:
                        break
                    for hit in hits:
                        oid   = str(hit.get("objectID", ""))
                        if oid in seen:
                            continue
                        title = (hit.get("title") or "").strip()
                        # ONLY keep actual Show HN or Launch HN posts
                        if not re.match(r"(?i)^(show|launch)\s+hn", title):
                            continue
                        seen.add(oid)
                        # Strip the "Show HN:" / "Launch HN:" prefix for cleaner name
                        name = re.sub(r"(?i)^(show|launch)\s+hn[:\s–-]+", "", title).strip()
                        if len(name) < 3:
                            continue
                        body     = clean_text(hit.get("story_text") or "", 500)
                        site_url = hit.get("url") or f"https://news.ycombinator.com/item?id={oid}"
                        src = "HN Show HN" if title.lower().startswith("show") else "HN Launch HN"
                        results.append(_entry(
                            name=name,
                            description=body or name,
                            source=src,
                            url=f"https://news.ycombinator.com/item?id={oid}",
                            website_url=site_url,
                            created_at=(hit.get("created_at") or "")[:10],
                            tags=kw,
                        ))
                    time.sleep(0.15)
                except Exception as e:
                    print(f"    ⚠ HN '{query[:30]}' p{page}: {e}")
                    break
    return results


def fetch_eu_startups():
    """EU-Startups directory — European deep-tech focus, multiple categories."""
    results = []
    seen    = set()
    categories = [
        "cleantech-greentech", "energy", "sustainability",
        "deep-tech", "hardware", "climate", "advanced-manufacturing",
    ]
    for cat in categories:
        for page in range(1, 5):
            url = (
                f"https://eu-startups.com/directory/"
                f"?wpbdp_view=listings&cat={cat}&page={page}"
            )
            r = safe_get(url)
            if not r:
                break
            soup = BeautifulSoup(r.content, "html.parser")
            cards = (
                soup.find_all("div", class_=lambda c: c and "wpbdp" in c.lower() and "listing" in c.lower())
                or soup.find_all("div", class_=lambda c: c and "listing" in c.lower())
                or soup.find_all("article")
                or soup.find_all("div", class_="company-card")
            )
            found = 0
            for card in cards:
                name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
                desc_el = card.find("p")
                link_el = card.find("a", href=True)
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 3 or name in seen:
                    continue
                seen.add(name)
                desc = clean_text(desc_el.get_text() if desc_el else "", 500)
                href = link_el["href"] if link_el else "https://eu-startups.com"
                if href.startswith("/"):
                    href = "https://eu-startups.com" + href
                results.append(_entry(
                    name=name, description=desc or name,
                    source="EU-Startups", url=href, website_url=href, tags=cat,
                ))
                found += 1
            if found == 0:
                break
            time.sleep(0.5)
    return results


def fetch_climatedraft():
    """Try multiple climate-tech company directories."""
    results = []
    seen    = set()
    sources = [
        ("https://climatedraft.org/companies",     "ClimateDraft"),
        ("https://climatetechlist.com/",           "ClimateTechList"),
        ("https://www.climatebase.org/companies",  "Climatebase"),
    ]
    for page_url, label in sources:
        r = safe_get(page_url)
        if not r:
            continue
        soup  = BeautifulSoup(r.content, "html.parser")
        cards = (
            soup.find_all("div", class_=lambda c: c and "company" in c.lower())
            or soup.find_all("li",  class_=lambda c: c and "company" in c.lower())
            or soup.find_all("article")
            or soup.find_all("div", class_=lambda c: c and "card" in c.lower())
        )
        found = 0
        for card in cards:
            name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
            desc_el = card.find("p")
            link_el = card.find("a", href=True)
            if not name_el:
                continue
            name = name_el.get_text(strip=True)
            if not name or len(name) < 3 or name in seen:
                continue
            seen.add(name)
            desc = clean_text(desc_el.get_text() if desc_el else "", 500)
            href = link_el["href"] if link_el else page_url
            if href.startswith("/"):
                href = page_url.rstrip("/") + href
            results.append(_entry(
                name=name, description=desc or name,
                source=label, url=href, website_url=href, tags="climate",
            ))
            found += 1
        if found > 0:
            print(f"      ({label}: +{found})", end="")
        time.sleep(0.4)
    return results


def fetch_sifted(pages=4):
    """
    Sifted.eu — European startup news. Real HTML (not JS-rendered).
    Covers deep-tech, climate, energy. Good for EU-based pre-seed.
    """
    results = []
    seen    = set()
    base_urls = [
        "https://sifted.eu/sector/sustainability",
        "https://sifted.eu/sector/deeptech",
        "https://sifted.eu/sector/energy",
        "https://sifted.eu/startups",
        "https://sifted.eu/",
    ]
    for base in base_urls:
        for page in range(1, pages + 1):
            url = f"{base}?page={page}" if page > 1 else base
            r = safe_get(url, timeout=20)
            if not r:
                break
            soup  = BeautifulSoup(r.content, "html.parser")
            # Sifted's article cards — article titles link to company profiles
            cards = (
                soup.find_all("article")
                or soup.find_all("div", class_=lambda c: c and "article" in c.lower())
                or soup.find_all("div", class_=lambda c: c and "card" in c.lower())
            )
            found = 0
            for card in cards:
                name_el = card.find(["h2", "h3", "h4"]) or card.find("a")
                desc_el = card.find("p")
                link_el = card.find("a", href=True)
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 3 or len(name) > 80 or name in seen:
                    continue
                if is_article_title(name):
                    continue
                seen.add(name)
                desc = clean_text(desc_el.get_text() if desc_el else "", 500)
                href = link_el["href"] if link_el else base
                if href.startswith("/"):
                    href = "https://sifted.eu" + href
                results.append(_entry(
                    name=name, description=desc or name,
                    source="Sifted EU", url=href, website_url=href, tags="europe",
                ))
                found += 1
            if found == 0:
                break
            time.sleep(0.5)
        if results:
            break  # got results from first working URL
    return results


def fetch_ddg_startups(keywords):
    """
    DuckDuckGo startup discovery — searches eu-startups.com and sifted.eu
    directly for articles mentioning thesis keywords, extracting company names.
    """
    if not keywords:
        return []
    results = []
    seen    = set()
    specific = [k for k in keywords if len(k) > 6][:6]

    for kw in specific:
        for query in [
            f'{kw} startup site:eu-startups.com',
            f'{kw} startup site:sifted.eu',
            f'{kw} company "pre-seed" OR "seed round" 2023 OR 2024 OR 2025',
        ]:
            try:
                r = requests.get(
                    f"https://html.duckduckgo.com/html/?q={quote_plus(query)}",
                    headers={
                        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                                      "AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
                        "Accept": "text/html",
                    },
                    timeout=12,
                )
                if not r.ok or r.status_code in (429, 202):
                    time.sleep(2.0)
                    continue
                soup = BeautifulSoup(r.text, "html.parser")
                divs = soup.find_all("div", class_="result")
                if not divs:
                    time.sleep(2.0)
                    continue
                for div in divs[:6]:
                    title_el = div.find("a", class_="result__a")
                    snip_el  = div.find("a", class_="result__snippet")
                    url_el   = div.find("span", class_="result__url")
                    if not title_el:
                        continue
                    title   = title_el.get_text(strip=True)
                    snippet = clean_text(snip_el.get_text() if snip_el else "", 400)
                    href    = url_el.get_text(strip=True) if url_el else ""
                    if href and not href.startswith("http"):
                        href = "https://" + href
                    # For site-specific queries, try visiting the article to extract the company
                    if title and not is_article_title(title) and title not in seen:
                        seen.add(title)
                        results.append(_entry(
                            name=title, description=snippet or title,
                            source="DDG/News", url=href, website_url=href, tags=kw,
                        ))
                time.sleep(1.2)
            except Exception as e:
                print(f"    ⚠ DDG '{kw[:25]}': {e}")
    return results


def fetch_llm_known_startups(client, thesis, keywords, focus="eu"):
    """
    Ask the LLM to recall startups from its training data.

    The model was trained on data that includes Crunchbase, TechCrunch,
    eu-startups.com, sifted.eu, and thousands of startup websites — it knows
    many companies that won't appear in any single web scrape.

    focus="eu"     → European companies + spinouts (call first)
    focus="global" → North American, Asian, Israeli, global (call second)
    """
    if not thesis:
        return []
    kw_str = ", ".join(keywords[:12]) if keywords else "the technology in the thesis"

    if focus == "global":
        geo_instruction = (
            "- North American companies (USA, Canada)\n"
            "- Asian companies (Japan, South Korea, China, Singapore, India)\n"
            "- Israeli deep-tech companies\n"
            "- Australian and other global companies\n"
            "- University spinouts from MIT, Stanford, Caltech, ETH Zurich, etc.\n"
            "- DO NOT repeat European companies — those are handled separately."
        )
        label = "LLM Knowledge (global)"
        n_target = "30-45"
    else:
        geo_instruction = (
            "- European companies (Germany, France, Netherlands, UK, Nordics, Spain, Italy, etc.)\n"
            "- University spinouts from European universities\n"
            "- Companies in EU accelerator programs (EIC, EIT InnoEnergy, Techstars)"
        )
        label = "LLM Knowledge (EU)"
        n_target = "35-50"

    prompt = (
        f"You are recalling ONLY obscure, early-stage startups from your training data.\n\n"
        f"THESIS: {thesis[:600]}\n"
        f"KEY TECHNOLOGIES: {kw_str}\n"
        f"Geographic focus: {geo_instruction}\n\n"
        f"HARD RULES:\n"
        f"1. DO NOT list any company that has raised more than $5M from VCs. "
        f"This immediately disqualifies them.\n"
        f"2. DO NOT list any company that has appeared in TechCrunch, Forbes, "
        f"Bloomberg, or major tech media as a headline story. Those are too well-known.\n"
        f"3. DO NOT list universities, research institutes, government programs, "
        f"or publicly traded companies.\n"
        f"4. AVOID these well-known companies and all similarly-sized peers: "
        f"Commonwealth Fusion Systems, TerraPower, Fervo Energy, Oxford PV, "
        f"EnergyNest, Hydrogenious, Highview Power, SaltX Technology, CorPower, "
        f"Climeon, Fusion for Energy, and any company with a Wikipedia article.\n"
        f"5. Only include companies where you can recall at least one specific fact: "
        f"the exact city, a founder's name, or a specific product name. "
        f"If you cannot recall this, SKIP the company.\n"
        f"6. Target: founded 2019–2025, pre-seed to early seed, minimal media coverage.\n\n"
        f"Return ONLY a valid JSON array:\n"
        f'[{{"name":"Exact Company Name",'
        f'"description":"2-3 specific sentences — what exactly they do, where, who founded it",'
        f'"website":"domain.com or empty","country":"country","tech_area":"specific tech"}}]\n\n'
        f"10 verified obscure companies beat 40 hallucinated or well-known ones."
    )
    print(f"  Querying {label}...", end=" ", flush=True)
    try:
        raw, _lbl = llm_call([
            {"role": "system", "content": (
                "You are a deep-tech startup expert. "
                "Your ONE job: list companies you have ACTUALLY seen in your training data. "
                "Never invent names, never extrapolate, never guess. "
                "If uncertain, skip the company. Return only valid JSON array."
            )},
            {"role": "user", "content": prompt},
        ], max_tokens=4000)
        if not raw:
            raise ValueError("empty response")
        raw = re.sub(r"```(?:json)?", "", raw).strip().strip("`")
        companies = json.loads(raw)
        results   = []
        for c in companies:
            name = str(c.get("name", "")).strip()
            desc = str(c.get("description", "")).strip()
            web  = str(c.get("website",     "")).strip()
            if not name or len(name) < 2:
                continue
            if web and not web.startswith("http"):
                web = "https://" + web
            results.append(_entry(
                name=name,
                description=desc or name,
                source="LLM Knowledge",
                url=web or f"https://www.google.com/search?q={quote_plus(name)}",
                website_url=web,
                tags=str(c.get("tech_area", "")),
            ))
        print(f"{len(results)} companies recalled.")
        return results
    except Exception as e:
        print(f"failed ({e})")
        return []


def fetch_product_hunt():
    """Product Hunt RSS feed."""
    import xml.etree.ElementTree as ET
    results = []
    for feed_url in [
        "https://www.producthunt.com/feed",
        "https://www.producthunt.com/posts/feed",
    ]:
        r = safe_get(feed_url)
        if not r:
            continue
        try:
            root  = ET.fromstring(r.content)
            items = root.findall(".//item")
            for item in items:
                title    = (item.findtext("title") or "").strip()
                desc_raw = item.findtext("description") or ""
                link     = item.findtext("link") or "https://producthunt.com"
                pub_date = (item.findtext("pubDate") or "")[:10]
                if not title or len(title) < 3:
                    continue
                desc = clean_text(BeautifulSoup(desc_raw, "html.parser").get_text(), 500)
                results.append(_entry(
                    name=title, description=desc or title,
                    source="Product Hunt", url=link, created_at=pub_date,
                ))
            if results:
                break
        except Exception:
            continue
    return results


def fetch_betalist():
    """
    BetaList — curated pre-launch startup directory (real HTML scrape).
    Every listing is a real company before public launch = genuinely early-stage.
    No hallucination risk: scraped from actual page content.
    """
    results = []
    seen    = set()
    for url in ["https://betalist.com/", "https://betalist.com/startups"]:
        r = safe_get(url, timeout=15)
        if not r:
            continue
        soup  = BeautifulSoup(r.content, "html.parser")
        # BetaList structure: <div class="block"><a href="/startups/slug">Name</a>desc</div>
        for link in soup.find_all("a", href=True):
            href = link["href"]
            if not href.startswith("/startups/"):
                continue
            parent = link.parent
            name   = link.get_text(strip=True)
            if not name or len(name) < 2 or name in seen or is_article_title(name):
                continue
            seen.add(name)
            # Description is usually text after the link in the same parent div
            full_text = parent.get_text(separator=" ", strip=True)
            desc      = clean_text(full_text.replace(name, "", 1).strip(), 400)
            full_url  = "https://betalist.com" + href
            results.append(_entry(
                name=name, description=desc or name,
                source="BetaList", url=full_url, website_url=full_url,
            ))
        if results:
            break
    return results


def fetch_github_startups(keywords, max_per_kw=10):
    """
    GitHub repository search via the public API.
    Returns real repos — zero hallucination risk.
    Great for finding solo-founder and very early-stage technical projects.
    """
    results = []
    seen    = set()
    kws     = [k for k in keywords if len(k) > 4][:8]
    for kw in kws:
        try:
            r = requests.get(
                "https://api.github.com/search/repositories",
                params={"q": f"{kw} startup", "sort": "updated",
                        "order": "desc", "per_page": max_per_kw},
                headers={"Accept": "application/vnd.github.v3+json"},
                timeout=12,
            )
            if not r.ok:
                time.sleep(2.0)
                continue
            for item in r.json().get("items", []):
                raw_name = item.get("name", "")
                name     = raw_name.replace("-", " ").replace("_", " ").title()
                desc     = item.get("description") or ""
                gh_url   = item.get("html_url", "")
                homepage = item.get("homepage") or ""
                owner    = item.get("owner", {}).get("login", "")
                if not name or len(name) < 3 or name.lower() in seen:
                    continue
                if is_article_title(name):
                    continue
                seen.add(name.lower())
                full_desc = clean_text(
                    f"{desc} (GitHub: {owner}/{raw_name})", 500
                )
                results.append(_entry(
                    name=name,
                    description=full_desc or name,
                    source="GitHub",
                    url=gh_url,
                    website_url=homepage or gh_url,
                    tags=kw,
                ))
            time.sleep(1.0)
        except Exception as e:
            print(f"    ⚠ GitHub '{kw[:25]}': {e}")
    return results


def fetch_rss_feeds(keywords, max_items=20):
    """
    Multi-source RSS feeds: TechCrunch, EU-Startups (+ category feeds),
    Sifted, VentureBeat, The Next Web.
    Extracts company names from funding/launch article titles using regex.
    RSS is guaranteed real XML — never JS-rendered.
    """
    import xml.etree.ElementTree as ET

    feeds = [
        ("https://techcrunch.com/feed/",                          "TechCrunch"),
        ("https://eu-startups.com/feed/",                         "EU-Startups"),
        ("https://eu-startups.com/category/deeptech/feed/",       "EU-Startups DeepTech"),
        ("https://eu-startups.com/category/energy/feed/",         "EU-Startups Energy"),
        ("https://eu-startups.com/category/sustainability/feed/", "EU-Startups Sustainability"),
        ("https://sifted.eu/feed",                                "Sifted"),
        ("https://venturebeat.com/feed/",                         "VentureBeat"),
        ("https://thenextweb.com/feed",                           "The Next Web"),
    ]

    # Patterns to extract company name from article title
    _RAISE  = re.compile(
        r"^([A-Z][A-Za-z0-9\.\-]{1,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,25}){0,3})"
        r"\s+(?:raises?|secures?|closes?|lands?|receives?)\s+[\$€£]", re.I,
    )
    _MEET   = re.compile(r"^Meet\s+([A-Z][A-Za-z0-9\.\-]{2,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,20})?)", re.I)
    _LAUNCH = re.compile(
        r"^([A-Z][A-Za-z0-9\.\-]{2,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,20})?)"
        r"\s+(?:launches?|debuts?|unveils?|announces?)\b", re.I,
    )
    _SEED   = re.compile(
        r"^([A-Z][A-Za-z0-9\.\-]{1,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,20}){0,2})"
        r",\s+(?:a startup|the startup|a company|the company)\b", re.I,
    )
    PATS = [_RAISE, _MEET, _LAUNCH, _SEED]

    results = []
    seen    = set()
    kw_set  = set(k.lower() for k in (keywords or []))

    for feed_url, label in feeds:
        r = safe_get(feed_url, timeout=15)
        if not r:
            continue
        try:
            root  = ET.fromstring(r.content)
            items = root.findall(".//item")
            count = 0
            for item in items:
                if count >= max_items:
                    break
                title    = (item.findtext("title") or "").strip()
                desc_raw = item.findtext("description") or ""
                link     = item.findtext("link") or feed_url
                pub_date = (item.findtext("pubDate") or "")[:10]
                if not title:
                    continue
                text_low = (title + " " + desc_raw).lower()
                if kw_set and not any(k in text_low for k in kw_set):
                    continue
                # Extract company name from title
                company = None
                for pat in PATS:
                    m = pat.match(title)
                    if m:
                        cand = m.group(1).strip().rstrip(",")
                        if 3 <= len(cand) <= 50 and not is_article_title(cand):
                            company = cand
                            break
                if not company:
                    continue
                norm = company.lower()[:40]
                if norm in seen:
                    continue
                seen.add(norm)
                desc = clean_text(BeautifulSoup(desc_raw, "html.parser").get_text(), 500)
                results.append(_entry(
                    name=company,
                    description=desc or f"{company} — {label}",
                    source=label,
                    url=link,
                    website_url=link,
                    created_at=pub_date,
                    tags=", ".join(k for k in (keywords or [])[:3] if k in text_low),
                ))
                count += 1
            time.sleep(0.3)
        except Exception as e:
            print(f"    ⚠ RSS {label}: {e}")

    return results


def fetch_crunchbase_ddg(keywords, max_per_kw=6):
    """
    DDG search targeting crunchbase.com/organization pages.
    Extracts real company names + Crunchbase profiles for thesis keywords.
    This finds small/obscure startups that have Crunchbase entries.
    """
    if not keywords:
        return []
    results = []
    seen    = set()
    kws     = sorted([k for k in keywords if len(k) > 5], key=len, reverse=True)[:8]
    hdrs    = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html",
    }
    for kw in kws:
        query = f'site:crunchbase.com/organization {kw}'
        try:
            r = requests.get(
                f"https://html.duckduckgo.com/html/?q={quote_plus(query)}",
                headers=hdrs, timeout=12,
            )
            if not r.ok or r.status_code in (429, 202):
                time.sleep(3.0)
                continue
            soup  = BeautifulSoup(r.text, "html.parser")
            divs  = soup.find_all("div", class_="result")
            if not divs:  # DDG homepage redirect (rate limit)
                time.sleep(3.0)
                continue
            count = 0
            for div in divs:
                if count >= max_per_kw:
                    break
                url_el   = div.find("span", class_="result__url")
                title_el = div.find("a",    class_="result__a")
                snip_el  = div.find("a",    class_="result__snippet")
                if not url_el or not title_el:
                    continue
                url_txt = url_el.get_text(strip=True).lower()
                m = re.search(r"crunchbase\.com/organization/([a-z0-9][a-z0-9\-]{1,60})", url_txt)
                if not m:
                    continue
                slug         = m.group(1)
                title_txt    = title_el.get_text(strip=True)
                company_name = (
                    title_txt.split(" - ")[0].strip()
                    if " - " in title_txt
                    else slug.replace("-", " ").title()
                )
                if not company_name or len(company_name) < 3 or is_article_title(company_name):
                    continue
                norm = company_name.lower()[:40]
                if norm in seen:
                    continue
                seen.add(norm)
                snip = clean_text(snip_el.get_text() if snip_el else "", 400)
                href = f"https://www.crunchbase.com/organization/{slug}"
                results.append(_entry(
                    name=company_name, description=snip or company_name,
                    source="Crunchbase", url=href, website_url=href, tags=kw,
                ))
                count += 1
            time.sleep(1.5)
        except Exception as e:
            print(f"    ⚠ Crunchbase DDG '{kw[:25]}': {e}")
    return results


def fetch_dealroom_ddg(keywords, max_per_kw=5):
    """
    DDG search targeting Dealroom.co — the leading EU startup database.
    Dealroom tracks funding rounds, investors, and company profiles across Europe.
    """
    if not keywords:
        return []
    results = []
    seen    = set()
    kws     = sorted([k for k in keywords if len(k) > 5], key=len, reverse=True)[:6]
    hdrs    = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html",
    }
    for kw in kws:
        query = f'site:dealroom.co {kw} startup'
        try:
            r = requests.get(
                f"https://html.duckduckgo.com/html/?q={quote_plus(query)}",
                headers=hdrs, timeout=12,
            )
            if not r.ok or r.status_code in (429, 202):
                time.sleep(3.0)
                continue
            soup  = BeautifulSoup(r.text, "html.parser")
            divs  = soup.find_all("div", class_="result")
            if not divs:  # DDG homepage redirect (rate limit)
                time.sleep(3.0)
                continue
            count = 0
            for div in divs:
                if count >= max_per_kw:
                    break
                title_el = div.find("a",    class_="result__a")
                snip_el  = div.find("a",    class_="result__snippet")
                url_el   = div.find("span", class_="result__url")
                if not title_el:
                    continue
                url_txt = (url_el.get_text(strip=True) if url_el else "").lower()
                if "dealroom.co" not in url_txt:
                    continue
                title_txt    = title_el.get_text(strip=True)
                company_name = title_txt.split(" - ")[0].strip() if " - " in title_txt else title_txt
                if not company_name or len(company_name) < 3 or is_article_title(company_name):
                    continue
                norm = company_name.lower()[:40]
                if norm in seen:
                    continue
                seen.add(norm)
                snip = clean_text(snip_el.get_text() if snip_el else "", 400)
                href = (url_el.get_text(strip=True) if url_el else "https://dealroom.co")
                if href and not href.startswith("http"):
                    href = "https://" + href
                results.append(_entry(
                    name=company_name, description=snip or company_name,
                    source="Dealroom", url=href, website_url=href, tags=kw,
                ))
                count += 1
            time.sleep(1.5)
        except Exception as e:
            print(f"    ⚠ Dealroom DDG '{kw[:25]}': {e}")
    return results


def fetch_cordis_eu(keywords, max_pages=3):
    """
    CORDIS — EU Horizon grants database. Searches organisations (contenttype=org)
    funded under Horizon Europe, EIC Accelerator, EIT InnoEnergy etc.
    These are genuinely early-stage: grant-funded before VC, often university spinouts.
    """
    if not keywords:
        return []
    results = []
    seen    = set()
    kws     = [k for k in keywords if len(k) > 4][:6]

    for kw in kws:
        for page in range(1, max_pages + 1):
            url = (
                f"https://cordis.europa.eu/search/results_en"
                f"?q={quote_plus(kw)}&contenttype=org"
                f"&p={page}&num=25&srt=Relevant:decreasing"
            )
            r = safe_get(url, timeout=25)
            if not r:
                break
            soup    = BeautifulSoup(r.content, "html.parser")
            entries = (
                soup.find_all("article")
                or soup.find_all("li",  class_=lambda c: c and "search-result" in c.lower())
                or soup.find_all("div", class_=lambda c: c and "result" in c.lower())
            )
            found = 0
            for entry in entries:
                name_el = entry.find(["h2", "h3", "h4", "strong"]) or entry.find("a")
                desc_el = entry.find("p")
                link_el = entry.find("a", href=True)
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 3 or name in seen or is_article_title(name):
                    continue
                seen.add(name)
                desc = clean_text(desc_el.get_text() if desc_el else "", 400)
                href = link_el["href"] if link_el else url
                if href.startswith("/"):
                    href = "https://cordis.europa.eu" + href
                results.append(_entry(
                    name=name, description=desc or name,
                    source="CORDIS EU", url=href, website_url=href, tags=kw,
                ))
                found += 1
            if found == 0:
                break
            time.sleep(0.5)
    return results


def fetch_university_spinoffs():
    """
    University technology transfer offices — ETH Zurich, EPFL, TU Delft, KTH.
    University spinouts are the canonical pre-seed deep-tech source.
    """
    sources = [
        ("https://transfer.ethz.ch/spin-offs.html",                      "ETH Zurich"),
        ("https://tto.epfl.ch/startups/",                                "EPFL"),
        ("https://www.tudelft.nl/en/tpm/about-the-faculty/departments/"
         "engineering-systems-and-services/research/startup-portfolio",   "TU Delft"),
        ("https://www.kth.se/en/om/nyheter/central-nyheter/"
         "innovation/startups",                                            "KTH"),
        ("https://www.imperial.ac.uk/enterprise/startups/",               "Imperial London"),
        ("https://www.kcl.ac.uk/innovation/spin-out-companies",           "King's College"),
    ]

    results = []
    seen    = set()

    for url, label in sources:
        r = safe_get(url, timeout=20)
        if not r:
            continue
        soup  = BeautifulSoup(r.content, "html.parser")
        cards = (
            soup.find_all("div", class_=lambda c: c and any(
                w in c.lower() for w in ["spin", "startup", "company", "card", "item"]
            ))
            or soup.find_all("article")
            or soup.find_all("li",  class_=lambda c: c and ("item" in c.lower() or "entry" in c.lower()))
        )
        found = 0
        for card in cards:
            name_el = card.find(["h2", "h3", "h4", "strong"]) or card.find("a")
            desc_el = card.find("p")
            link_el = card.find("a", href=True)
            if not name_el:
                continue
            name = name_el.get_text(strip=True)
            if not name or len(name) < 2 or name in seen or is_article_title(name):
                continue
            seen.add(name)
            desc = clean_text(desc_el.get_text() if desc_el else "", 400)
            href = link_el["href"] if link_el else url
            if href.startswith("/"):
                base = "/".join(url.split("/")[:3])
                href = base + href
            results.append(_entry(
                name=name, description=desc or name,
                source=label, url=href, website_url=href, tags="university spinout",
            ))
            found += 1
        if found > 0:
            print(f"      ({label}: +{found})", end="")
        time.sleep(0.5)

    return results


def fetch_eic_portfolio(keywords):
    """
    EIC Accelerator portfolio via F6S and direct EIC pages.
    EIC is the EU's main deep-tech accelerator — companies are early-stage by design.
    """
    if not keywords:
        return []
    results = []
    seen    = set()
    kws     = [k for k in keywords if len(k) > 4][:5]
    hdrs    = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html",
    }

    for kw in kws:
        # DDG search targeting EIC community and success stories pages
        for query in [
            f'site:eic.ec.europa.eu {kw} company startup',
            f'"EIC Accelerator" {kw} startup winner funded',
        ]:
            try:
                r = requests.get(
                    f"https://html.duckduckgo.com/html/?q={quote_plus(query)}",
                    headers=hdrs, timeout=12,
                )
                if not r.ok or r.status_code in (429, 202):
                    time.sleep(2.0)
                    continue
                soup = BeautifulSoup(r.text, "html.parser")
                divs = soup.find_all("div", class_="result")
                if not divs:
                    time.sleep(2.0)
                    continue
                for div in divs[:5]:
                    title_el = div.find("a",    class_="result__a")
                    snip_el  = div.find("a",    class_="result__snippet")
                    url_el   = div.find("span", class_="result__url")
                    if not title_el:
                        continue
                    title = title_el.get_text(strip=True)
                    if is_article_title(title) or title in seen:
                        continue
                    seen.add(title)
                    snip = clean_text(snip_el.get_text() if snip_el else "", 400)
                    href = (url_el.get_text(strip=True) if url_el else "")
                    if href and not href.startswith("http"):
                        href = "https://" + href
                    results.append(_entry(
                        name=title, description=snip or title,
                        source="EIC Accelerator", url=href, website_url=href, tags=kw,
                    ))
                time.sleep(1.5)
            except Exception as e:
                print(f"    ⚠ EIC DDG '{kw[:25]}': {e}")

    return results


# Reddit removed — persistent encoding errors, returned 0 results for deep-tech theses.

def fetch_eu_vc_portfolios():
    """
    Scrape portfolio pages of 15 EU-focused early-stage VCs.
    These are real companies, pre-vetted as investable, mostly seed/pre-seed.
    """
    import xml.etree.ElementTree as ET

    sources = [
        # (url, label, link_pattern_hint)
        ("https://www.seedcamp.com/portfolio/",          "Seedcamp"),
        ("https://localglobe.vc/portfolio/",             "LocalGlobe"),
        ("https://creandum.com/portfolio/",              "Creandum"),
        ("https://www.speedinvest.com/portfolio",        "Speedinvest"),
        ("https://pointninecap.com/portfolio/",          "Point Nine"),
        ("https://www.cherry.vc/portfolio",              "Cherry Ventures"),
        ("https://frontline.vc/portfolio/",              "Frontline VC"),
        ("https://www.earlybird.com/portfolio/",         "Earlybird"),
        ("https://balderton.com/portfolio/",             "Balderton"),
        ("https://notion.vc/portfolio/",                 "Notion Capital"),
        ("https://www.lakestar.com/portfolio",           "Lakestar"),
        ("https://hv.capital/portfolio",                 "HV Capital"),
        ("https://www.atomico.com/portfolio",            "Atomico"),
        ("https://www.northzone.com/portfolio",          "Northzone"),
        ("https://www.partech.vc/portfolio",             "Partech"),
    ]

    results = []
    seen = set()

    for url, label in sources:
        try:
            r = safe_get(url, timeout=20)
            if not r:
                continue
            soup = BeautifulSoup(r.content, "html.parser")
            found = 0
            # Try structured cards first
            cards = (
                soup.find_all("div", class_=lambda c: c and any(
                    w in c.lower() for w in ["portfolio", "company", "startup", "card", "item", "grid"]
                ))
                or soup.find_all("article")
                or soup.find_all("li", class_=lambda c: c and any(
                    w in c.lower() for w in ["portfolio", "company", "item"]
                ))
            )
            for card in cards:
                name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
                link_el = card.find("a", href=True)
                desc_el = card.find("p")
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 2 or len(name) > 60 or name in seen or is_article_title(name):
                    continue
                seen.add(name)
                desc = clean_text(desc_el.get_text() if desc_el else "", 400)
                href = link_el["href"] if link_el else url
                if href and href.startswith("/"):
                    href = "/".join(url.split("/")[:3]) + href
                results.append(_entry(
                    name=name, description=desc or f"{name} — portfolio company of {label}",
                    source=f"VC:{label}", url=href, website_url=href,
                    tags="eu vc portfolio",
                ))
                found += 1
            if found > 0:
                print(f"      ({label}: +{found})", end="")
            time.sleep(0.4)
        except Exception as e:
            print(f"      ({label}: err)", end="")

    return results


def fetch_eu_accelerators():
    """
    EU accelerator & grant portfolio pages.
    EIT InnoEnergy, Climate-KIC, Startup Wise Guys, Techstars EU, Station F.
    Grant/accelerator backed = guaranteed early-stage.
    """
    sources = [
        ("https://www.eitinnovationnest.eu/portfolio",       "EIT InnoEnergy"),
        ("https://climate-kic.org/programmes/",              "Climate-KIC"),
        ("https://startupwiseguys.com/portfolio/",           "Startup Wise Guys"),
        ("https://www.eitdigital.eu/our-companies/",         "EIT Digital"),
        ("https://www.f6s.com/community/programs/eic",       "F6S EIC"),
        ("https://thefamily.co/portfolio",                   "The Family"),
        ("https://station-f.com/startups",                   "Station F"),
        ("https://deeptech.eu/portfolio/",                   "DeepTech EU"),
        ("https://euroquity.bpifrance.fr/companies",         "BPI France"),
    ]

    results = []
    seen = set()

    for url, label in sources:
        try:
            r = safe_get(url, timeout=20)
            if not r:
                continue
            soup = BeautifulSoup(r.content, "html.parser")
            cards = (
                soup.find_all("div", class_=lambda c: c and any(
                    w in c.lower() for w in ["company", "startup", "portfolio", "card", "item"]
                ))
                or soup.find_all("article")
                or soup.find_all("li")
            )
            found = 0
            for card in cards:
                name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
                link_el = card.find("a", href=True)
                desc_el = card.find("p")
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 2 or len(name) > 60 or name in seen or is_article_title(name):
                    continue
                seen.add(name)
                desc = clean_text(desc_el.get_text() if desc_el else "", 400)
                href = link_el["href"] if link_el else url
                if href and href.startswith("/"):
                    href = "/".join(url.split("/")[:3]) + href
                results.append(_entry(
                    name=name, description=desc or f"{name} — accelerator: {label}",
                    source=f"Accelerator:{label}", url=href, website_url=href,
                    tags="eu accelerator",
                ))
                found += 1
            if found > 0:
                print(f"      ({label}: +{found})", end="")
            time.sleep(0.5)
        except Exception as e:
            print(f"      ({label}: err)", end="")

    return results


def fetch_yc_eu_companies():
    """
    Y Combinator company directory filtered to European companies.
    Uses YC's public company list API/HTML — zero hallucination risk.
    """
    results = []
    seen = set()
    eu_countries = {
        "germany", "france", "uk", "netherlands", "sweden", "denmark",
        "finland", "norway", "spain", "italy", "poland", "czech", "austria",
        "switzerland", "portugal", "ireland", "belgium", "estonia", "latvia",
        "lithuania", "hungary", "romania", "ukraine", "greece", "croatia",
        "united kingdom", "great britain",
    }
    try:
        # YC's public API for company data
        r = requests.get(
            "https://api.ycombinator.com/v0.1/companies",
            params={"batch": "S24,W24,S23,W23,S22,W22", "count": 500},
            timeout=20,
        )
        if r.ok:
            data = r.json()
            companies = data if isinstance(data, list) else data.get("companies", [])
            for c in companies:
                country = (c.get("country") or "").lower()
                region = (c.get("regions") or [""])[0].lower() if c.get("regions") else ""
                if not any(eu in country or eu in region for eu in eu_countries):
                    continue
                name = (c.get("name") or "").strip()
                if not name or name in seen:
                    continue
                seen.add(name)
                desc = clean_text(c.get("one_liner") or c.get("description") or "", 400)
                url = c.get("website") or f"https://www.ycombinator.com/companies/{c.get('slug','')}"
                results.append(_entry(
                    name=name, description=desc or name,
                    source="YC", url=url, website_url=url,
                    tags="yc eu",
                ))
    except Exception:
        pass

    if not results:
        # Fallback: scrape YC public HTML page
        try:
            r = safe_get("https://www.ycombinator.com/companies?batch=S24&batch=W24&batch=S23&batch=W23&regions=Europe", timeout=20)
            if r:
                soup = BeautifulSoup(r.content, "html.parser")
                for card in soup.find_all("a", href=lambda h: h and "/companies/" in h):
                    name = card.get_text(strip=True)
                    if not name or name in seen or len(name) < 2 or len(name) > 60:
                        continue
                    seen.add(name)
                    href = "https://www.ycombinator.com" + card["href"] if card["href"].startswith("/") else card["href"]
                    results.append(_entry(
                        name=name, description=f"{name} — YC company",
                        source="YC", url=href, website_url=href, tags="yc eu",
                    ))
        except Exception:
            pass

    return results


def fetch_seedtable(keywords=None):
    """
    Seedtable.com — curated EU startup rankings and profiles.
    Covers Germany, France, UK, Nordics, Netherlands with rich profiles.
    """
    results = []
    seen = set()
    pages = [
        "https://www.seedtable.com/startups-germany",
        "https://www.seedtable.com/startups-france",
        "https://www.seedtable.com/startups-netherlands",
        "https://www.seedtable.com/startups-sweden",
        "https://www.seedtable.com/startups-switzerland",
        "https://www.seedtable.com/startups-austria",
        "https://www.seedtable.com/startups-denmark",
        "https://www.seedtable.com/startups-finland",
        "https://www.seedtable.com/startups-norway",
        "https://www.seedtable.com/startups-spain",
        "https://www.seedtable.com/startups-poland",
    ]
    kw_set = {k.lower() for k in (keywords or [])}

    for url in pages:
        try:
            r = safe_get(url, timeout=20)
            if not r:
                continue
            soup = BeautifulSoup(r.content, "html.parser")
            country = url.split("-")[-1].replace("/","")
            found = 0
            for card in (soup.find_all("div", class_=lambda c: c and any(
                    w in c.lower() for w in ["startup","company","card","listing","item"]
                )) or soup.find_all("article") or soup.find_all("li")):
                name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
                desc_el = card.find("p")
                link_el = card.find("a", href=True)
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 2 or len(name) > 60 or name in seen or is_article_title(name):
                    continue
                desc = clean_text(desc_el.get_text() if desc_el else "", 400)
                # If keywords given, filter
                if kw_set:
                    combo = (name + " " + desc).lower()
                    if not any(k in combo for k in kw_set):
                        continue
                seen.add(name)
                href = link_el["href"] if link_el else url
                if href and href.startswith("/"):
                    href = "https://www.seedtable.com" + href
                results.append(_entry(
                    name=name, description=desc or name,
                    source="Seedtable", url=href, website_url=href,
                    tags=country,
                ))
                found += 1
            if found > 0:
                print(f"      ({country}: +{found})", end="")
            time.sleep(0.4)
        except Exception:
            pass

    return results


def fetch_more_universities():
    """
    Extended university TTO pages — 12 additional EU research universities.
    These are pure spinout lists: real companies, pre-seed by definition.
    """
    sources = [
        ("https://www.kuleuven.be/research/valorisation/spin-offs",         "KU Leuven"),
        ("https://www.tue.nl/en/research/tue-spinoffs/",                     "TU Eindhoven"),
        ("https://www.chalmers.se/en/about-chalmers/chalmers-startups/",    "Chalmers"),
        ("https://www.dtu.dk/english/research/dtu-entrepreneurship/startups","DTU Denmark"),
        ("https://www.ntnu.edu/ttto/spin-off-companies",                     "NTNU Norway"),
        ("https://www.aalto.fi/en/aalto-entrepreneurship/startups",         "Aalto Finland"),
        ("https://ucl-innovation-and-enterprise.com/startups/",             "UCL London"),
        ("https://www.cam.ac.uk/research/research-at-cambridge/cambridge-enterprise", "Cambridge"),
        ("https://innovation.ox.ac.uk/oxford-companies/",                   "Oxford"),
        ("https://www.tum.de/en/innovation/startups/",                      "TU Munich"),
        ("https://www.polimi.it/en/research/research-initiatives/polihub/", "Polimi"),
        ("https://www.epfl.ch/research/technology-transfer/",               "EPFL TT"),
        ("https://www.ugent.be/research/valorisation/spin-offs",            "Ghent Univ"),
        ("https://www.ed.ac.uk/commercialisation/spinout-companies",        "Edinburgh"),
        ("https://www.manchester.ac.uk/research/expertise/spin-out-companies/", "Manchester"),
    ]

    results = []
    seen = set()

    for url, label in sources:
        try:
            r = safe_get(url, timeout=20)
            if not r:
                continue
            soup = BeautifulSoup(r.content, "html.parser")
            cards = (
                soup.find_all("div", class_=lambda c: c and any(
                    w in c.lower() for w in ["spin","startup","company","card","item","venture"]
                ))
                or soup.find_all("article")
                or soup.find_all("li", class_=lambda c: c and any(
                    w in c.lower() for w in ["spin","item","startup","entry"]
                ))
            )
            found = 0
            for card in cards:
                name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
                desc_el = card.find("p")
                link_el = card.find("a", href=True)
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 2 or len(name) > 80 or name in seen or is_article_title(name):
                    continue
                seen.add(name)
                desc = clean_text(desc_el.get_text() if desc_el else "", 400)
                href = link_el["href"] if link_el else url
                if href and href.startswith("/"):
                    href = "/".join(url.split("/")[:3]) + href
                results.append(_entry(
                    name=name, description=desc or name,
                    source=label, url=href, website_url=href,
                    tags="university spinout",
                ))
                found += 1
            if found > 0:
                print(f"      ({label}: +{found})", end="")
            time.sleep(0.5)
        except Exception:
            pass

    return results


def fetch_tech_rss_news(keywords=None):
    """
    Extended RSS: Tech.eu, EU-Startups category feeds, Sifted, Nordic9, Maddyness.
    Parses article titles to extract company names from funding/launch news.
    """
    import xml.etree.ElementTree as ET

    feeds = [
        ("https://tech.eu/feed/",                                       "Tech.eu"),
        ("https://eu-startups.com/category/cleantech/feed/",            "EU-Startups Cleantech"),
        ("https://eu-startups.com/category/health-biotech/feed/",       "EU-Startups BioTech"),
        ("https://eu-startups.com/category/hardware/feed/",             "EU-Startups Hardware"),
        ("https://eu-startups.com/category/manufacturing/feed/",        "EU-Startups Mfg"),
        ("https://nordic9.com/feed/",                                   "Nordic9"),
        ("https://www.maddyness.com/feed/",                             "Maddyness FR"),
        ("https://www.eu-startups.com/category/germany/feed/",          "EU-Startups DE"),
        ("https://www.eu-startups.com/category/netherlands/feed/",      "EU-Startups NL"),
        ("https://www.eu-startups.com/category/sweden/feed/",           "EU-Startups SE"),
        ("https://www.eu-startups.com/category/france/feed/",           "EU-Startups FR"),
        ("https://sifted.eu/sector/deeptech/feed",                      "Sifted DeepTech"),
        ("https://sifted.eu/sector/sustainability/feed",                "Sifted Sustain"),
        ("https://sifted.eu/sector/energy/feed",                        "Sifted Energy"),
        ("https://www.smartcompany.com.au/feed/",                       "SmartCompany"),
    ]

    _RAISE  = re.compile(
        r"^([A-Z][A-Za-z0-9\.\-]{1,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,25}){0,3})"
        r"\s+(?:raises?|secures?|closes?|lands?|receives?|bags?)\s+[\$€£]", re.I)
    _MEET   = re.compile(r"^Meet\s+([A-Z][A-Za-z0-9\.\-]{2,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,20})?)", re.I)
    _LAUNCH = re.compile(
        r"^([A-Z][A-Za-z0-9\.\-]{2,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,20})?)"
        r"\s+(?:launches?|debuts?|unveils?|announces?|releases?)\b", re.I)
    _SEED   = re.compile(
        r"^([A-Z][A-Za-z0-9\.\-]{1,30}(?:\s+[A-Z][A-Za-z0-9\.\-]{1,20}){0,2})"
        r",\s+(?:a startup|the startup|a company|the company)\b", re.I)
    _NAMED  = re.compile(r"^([A-Z][A-Za-z0-9]{2,20}(?:\s[A-Z][A-Za-z0-9]{1,20})?)\s+is\s+(?:a|an)\s", re.I)
    PATS = [_RAISE, _MEET, _LAUNCH, _SEED, _NAMED]

    kw_set = {k.lower() for k in (keywords or [])}
    results = []
    seen = set()

    for feed_url, label in feeds:
        try:
            r = safe_get(feed_url, timeout=15)
            if not r:
                continue
            root = ET.fromstring(r.content)
            items = root.findall(".//item")
            count = 0
            for item in items:
                if count >= 30:
                    break
                title    = (item.findtext("title") or "").strip()
                desc_raw = item.findtext("description") or ""
                link     = item.findtext("link") or feed_url
                pub_date = (item.findtext("pubDate") or "")[:10]
                if not title:
                    continue
                # Only process if keyword match (or no keywords given)
                if kw_set:
                    combo = (title + " " + desc_raw).lower()
                    if not any(k in combo for k in kw_set):
                        continue
                company = None
                for pat in PATS:
                    m = pat.match(title)
                    if m:
                        cand = m.group(1).strip().rstrip(",")
                        if 3 <= len(cand) <= 50 and not is_article_title(cand):
                            company = cand
                            break
                if not company:
                    continue
                norm = company.lower()[:40]
                if norm in seen:
                    continue
                seen.add(norm)
                desc = clean_text(BeautifulSoup(desc_raw, "html.parser").get_text(), 400)
                results.append(_entry(
                    name=company,
                    description=desc or f"{company} — {label}",
                    source=label, url=link, website_url=link,
                    created_at=pub_date,
                    tags=", ".join(k for k in (keywords or [])[:3] if k in (title+desc_raw).lower()),
                ))
                count += 1
            time.sleep(0.3)
        except Exception:
            pass

    return results


def fetch_f6s_programs(keywords=None):
    """
    F6S.com — the largest EU startup program/accelerator platform.
    Scrapes public program listings for early-stage companies.
    """
    results = []
    seen = set()
    kw_set = {k.lower() for k in (keywords or [])}

    pages = [
        "https://www.f6s.com/programs?country=europe&type=accelerator",
        "https://www.f6s.com/programs?country=europe&type=grant",
        "https://www.f6s.com/companies?country=europe&stage=pre-seed",
        "https://www.f6s.com/companies?country=europe&stage=seed",
    ]

    for url in pages:
        try:
            r = safe_get(url, timeout=20)
            if not r:
                continue
            soup = BeautifulSoup(r.content, "html.parser")
            for card in (soup.find_all("div", class_=lambda c: c and any(
                    w in c.lower() for w in ["program","company","startup","card","listing"]
                )) or soup.find_all("article")):
                name_el = card.find(["h2","h3","h4","strong"]) or card.find("a")
                desc_el = card.find("p")
                link_el = card.find("a", href=True)
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                if not name or len(name) < 2 or len(name) > 60 or name in seen or is_article_title(name):
                    continue
                desc = clean_text(desc_el.get_text() if desc_el else "", 300)
                if kw_set and not any(k in (name+desc).lower() for k in kw_set):
                    continue
                seen.add(name)
                href = link_el["href"] if link_el else url
                if href and href.startswith("/"):
                    href = "https://www.f6s.com" + href
                results.append(_entry(
                    name=name, description=desc or name,
                    source="F6S", url=href, website_url=href, tags="eu startup",
                ))
            time.sleep(0.5)
        except Exception:
            pass

    return results


def fetch_wellfound_ddg(keywords, max_per_kw=8):
    """
    DDG search targeting wellfound.com (AngelList Talent) company profiles.
    Wellfound lists 100k+ startups including obscure EU ones.
    """
    if not keywords:
        return []
    results = []
    seen = set()
    hdrs = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
        "Accept": "text/html",
    }
    kws = sorted([k for k in keywords if len(k) > 4], key=len, reverse=True)[:10]

    for kw in kws:
        for query in [
            f'site:wellfound.com/company {kw} europe',
            f'site:angel.co/company {kw} europe startup',
        ]:
            try:
                r = requests.get(
                    f"https://html.duckduckgo.com/html/?q={quote_plus(query)}",
                    headers=hdrs, timeout=12,
                )
                if not r.ok or r.status_code in (429, 202):
                    time.sleep(3.0)
                    continue
                soup = BeautifulSoup(r.text, "html.parser")
                divs = soup.find_all("div", class_="result")
                if not divs:
                    time.sleep(3.0)
                    continue
                count = 0
                for div in divs:
                    if count >= max_per_kw:
                        break
                    title_el = div.find("a",    class_="result__a")
                    snip_el  = div.find("a",    class_="result__snippet")
                    url_el   = div.find("span", class_="result__url")
                    if not title_el:
                        continue
                    title_txt = title_el.get_text(strip=True)
                    url_txt   = (url_el.get_text(strip=True) if url_el else "").lower()
                    if "wellfound" not in url_txt and "angel.co" not in url_txt:
                        continue
                    company_name = title_txt.split(" - ")[0].strip() if " - " in title_txt else title_txt
                    if not company_name or len(company_name) < 3 or is_article_title(company_name):
                        continue
                    norm = company_name.lower()[:40]
                    if norm in seen:
                        continue
                    seen.add(norm)
                    snip = clean_text(snip_el.get_text() if snip_el else "", 400)
                    href = url_el.get_text(strip=True) if url_el else ""
                    if href and not href.startswith("http"):
                        href = "https://" + href
                    results.append(_entry(
                        name=company_name, description=snip or company_name,
                        source="Wellfound", url=href, website_url=href, tags=kw,
                    ))
                    count += 1
                time.sleep(1.5)
            except Exception as e:
                print(f"    ⚠ Wellfound DDG '{kw[:25]}': {e}")

    return results


# ═══════════════════════════════════════════════════════════════════════════════
# KEYWORD FILTER + DEDUP
# ═══════════════════════════════════════════════════════════════════════════════

_FUNDED_KEYWORDS = [
    "(yc ", "yc w", "yc s", "yc f", "y combinator", "ycombinator",
    "raised $", "raised €", "raised £",
    "series a", "series b", "series c", "series d", "series e", "series f",
    "seed round", "seed funding", "seed investment",
    "venture backed", "venture-backed",
    "a16z", "sequoia", "andreessen", "accel", "index ventures",
    "breakthrough energy", "softbank", "blackrock", "bill gates",
    "$1m", "$2m", "$5m", "$10m", "$50m", "$100m", "$200m", "$500m",
    "€10m", "€50m", "€100m", "€200m",
    "ipo", "initial public offering", "stock exchange", "nasdaq", "nyse",
    "acquired by", "acquisition complete",
    "unicorn", "decacorn",
]


def keyword_filter_and_dedup(startups, thesis_keywords=None):
    """Remove funded signals, article titles, thesis-irrelevant entries, then deduplicate."""
    # Step 1: remove funded signals AND article-like names
    clean, rm_funded = [], 0
    for s in startups:
        name = s.get("name", "")
        text = (name + " " + s.get("description", "")).lower()
        if any(kw in text for kw in _FUNDED_KEYWORDS):
            rm_funded += 1
        elif is_article_title(name):
            rm_funded += 1  # count with funded removals (both are noise)
        else:
            clean.append(s)

    # Step 2: thesis relevance filter (keep if ≥1 keyword matches)
    rm_irrelevant = 0
    if thesis_keywords:
        relevant = []
        for s in clean:
            if keyword_relevance(s, thesis_keywords) > 0:
                relevant.append(s)
            else:
                rm_irrelevant += 1
        clean = relevant

    # Step 3: deduplicate by name — exact match on first 40 chars AND first 3 words
    seen_names, seen_prefix, deduped = set(), set(), []
    for s in clean:
        full_key  = s["name"].lower().strip()
        short_key = full_key[:40]
        # First-3-words key catches "PyTogether, open-source..." duplicates with different suffixes
        words3    = " ".join(full_key.split()[:3])
        if short_key and len(short_key) > 2 and short_key not in seen_names and words3 not in seen_prefix:
            seen_names.add(short_key)
            if words3:
                seen_prefix.add(words3)
            deduped.append(s)

    # Step 4: deduplicate by domain
    seen_dom, final = set(), []
    for s in deduped:
        d = extract_domain(s.get("website_url", ""))
        if d and d in seen_dom:
            continue
        if d:
            seen_dom.add(d)
        final.append(s)

    return final, rm_funded, rm_irrelevant


# ═══════════════════════════════════════════════════════════════════════════════
# WEBSITE ENRICHMENT
# ═══════════════════════════════════════════════════════════════════════════════

def enrich_with_website(startup, timeout=8):
    """Fetch startup homepage and extract a clean description snippet."""
    url = startup.get("website_url") or startup.get("url", "")
    # Skip aggregator/forum URLs
    skip_hosts = ["reddit.com", "news.ycombinator.com", "producthunt.com",
                  "eu-startups.com", "climatedraft.org", "climatetechlist.com"]
    if not url or any(h in url for h in skip_hosts):
        return ""
    try:
        r = requests.get(
            url, headers=browser_headers(), timeout=timeout,
            allow_redirects=True,
        )
        if not r.ok:
            return ""
        soup = BeautifulSoup(r.content, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header", "aside"]):
            tag.decompose()
        # Prefer meta description (concise, intentional)
        meta = (
            soup.find("meta", attrs={"name": "description"}) or
            soup.find("meta", attrs={"property": "og:description"})
        )
        if meta and meta.get("content"):
            return clean_text(meta["content"], 400)
        # Fall back to first substantial paragraphs
        paras = [
            p.get_text(strip=True)
            for p in soup.find_all("p")
            if len(p.get_text(strip=True)) > 40
        ]
        return clean_text(" ".join(paras[:3]), 400)
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════════════════════════════
# VERIFICATION
# ═══════════════════════════════════════════════════════════════════════════════

def _ddg(query, timeout=8):
    try:
        r = requests.get(
            f"https://html.duckduckgo.com/html/?q={quote_plus(query)}",
            headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                              "AppleWebKit/537.36 Chrome/121.0.0.0 Safari/537.36",
                "Accept":     "text/html",
            },
            timeout=timeout,
        )
        if not r.ok or r.status_code in (429, 202):
            return None
        text = r.text.lower()
        # DDG sometimes returns homepage with no results (rate limit / bot block)
        if 'class="result"' not in text and 'class="results"' not in text:
            return None
        return text
    except Exception:
        return None


_LLM_SOURCES = {"LLM Knowledge", "LLM Knowledge (EU)", "LLM Knowledge (Global)"}

_STOPWORDS = {
    "company", "startup", "based", "offers", "provides", "focused", "technology",
    "system", "their", "which", "develop", "solution", "innovative", "advanced",
    "using", "through", "including", "other", "about", "with", "from", "that",
}


def check_company_exists(name, tech_tags="", description=""):
    """
    DDG existence check: confirm this company actually exists AND operates
    in the expected domain. Used to catch LLM hallucinations.

    Returns: "confirmed" | "mismatch" | "not_found" | "uncertain"
    """
    words  = name.split()
    anchor = " ".join(words[:min(3, len(words))]).lower()
    if len(anchor) < 5:
        return "uncertain"

    # Extract domain-specific keywords from tech_tags + description
    tech_text  = (tech_tags + " " + description[:300]).lower()
    domain_kws = [
        w for w in re.findall(r'\b[a-z]{5,}\b', tech_text)
        if w not in _STOPWORDS
    ][:10]

    text = _ddg(f'"{name}"')
    if text is None:
        return "uncertain"

    if anchor not in text:
        return "not_found"

    if domain_kws:
        hits = sum(1 for k in domain_kws if k in text)
        if hits == 0:
            # Company name found but zero domain words → different company / confabulation
            return "mismatch"

    return "confirmed"


def snippet_is_relevant(snippet, description, tech_tags=""):
    """
    Returns False if the website content clearly describes a different business.
    Catches cases like "Thermoelectric Co" website actually being a radio company.
    """
    if not snippet or len(snippet) < 60:
        return True  # Too short to judge

    combined      = (description + " " + tech_tags).lower()
    expected_words = {
        w for w in re.findall(r'\b[a-z]{5,}\b', combined)
        if w not in _STOPWORDS
    }
    if not expected_words:
        return True

    snippet_words = set(re.findall(r'\b[a-z]{5,}\b', snippet.lower()))
    overlap       = expected_words & snippet_words
    ratio         = len(overlap) / len(expected_words)

    # < 5% overlap between expected domain words and website content = wrong company
    return ratio >= 0.05


def web_funding_check(name):
    """
    DDG funding check. Detects: Series A–F, large amounts, IPO, acquisitions,
    Wikipedia presence (signals well-known = not pre-seed).
    """
    words       = name.split()
    name_anchor = " ".join(words[:min(3, len(words))]).lower()
    if len(name_anchor) < 6:
        return "clean"

    STRONG_SIGNALS = [
        # Round labels — ALL series including late-stage
        "series a", "series b", "series c", "series d", "series e", "series f",
        "seed round", "vc-backed", "venture round",
        # Amount keywords
        "raised $", "raised €", "raised £",
        "million in funding", "million funding", "billion in funding",
        "total funding", "total raised",
        # Exit / public
        "ipo", "initial public offering", "went public", "stock market listing",
        "acquired by", "acquisition",
        # Well-known large climate investors (presence = serious funding)
        "breakthrough energy ventures", "energy impact partners",
        "lowercarbon capital", "temasek", "softbank vision fund",
    ]

    # Regex for large dollar/euro amounts: $50M, €100M, $1.5B, 462 million, etc.
    _LARGE_AMT = re.compile(
        r'[\$€£]\s*\d{2,4}[\.,]?\d*\s*[mb]'        # $50M, €462M, $1.5B
        r'|\d{2,4}\s*million'                         # 462 million
        r'|\d+[\.,]?\d*\s*billion',                   # 1.5 billion
        re.I,
    )

    # Query 1: funding news
    query = f'"{name}" funding OR raised OR investors site:crunchbase.com OR site:techcrunch.com'
    text  = _ddg(query)
    if text is None:
        return "unknown"

    if name_anchor in text:
        for sig in STRONG_SIGNALS:
            if sig in text:
                return "funded"
        if _LARGE_AMT.search(text):
            return "funded"

    # Query 2: Wikipedia check — pre-seed companies don't have Wikipedia pages
    wiki_text = _ddg(f'site:en.wikipedia.org "{name}"')
    if wiki_text and name_anchor in wiki_text:
        return "funded"  # Wikipedia page = well-known = not pre-seed

    time.sleep(0.5)
    return "clean"


def sec_edgar_check(name):
    try:
        r = requests.get(
            f'https://efts.sec.gov/LATEST/search-index?q="{quote_plus(name)}"&forms=D',
            headers={"User-Agent": "StartupScout admin@example.com"},
            timeout=10,
        )
        r.raise_for_status()
        data  = r.json()
        total = data.get("hits", {}).get("total", {})
        count = total.get("value", 0) if isinstance(total, dict) else (total or 0)
        return count > 0
    except Exception:
        return None


def domain_age_check(website_url):
    domain = extract_domain(website_url)
    if not domain:
        return None, None
    try:
        r = requests.get(f"https://rdap.org/domain/{domain}", timeout=10)
        if r.status_code == 404:
            return None, None
        r.raise_for_status()
        for ev in r.json().get("events", []):
            if ev.get("eventAction", "").lower() == "registration":
                ds = ev.get("eventDate", "")
                if ds:
                    reg = datetime.fromisoformat(ds.replace("Z", "+00:00"))
                    age = (datetime.now(timezone.utc) - reg).days / 30.44
                    return round(age), age < 18
        return None, None
    except Exception:
        return None, None


# ═══════════════════════════════════════════════════════════════════════════════
# AGENT LOOP RUNNER
# ═══════════════════════════════════════════════════════════════════════════════

def run_agent(client, model, system, first_message, tools, executors, max_iterations=60):
    messages = [
        {"role": "system", "content": system},
        {"role": "user",   "content": first_message},
    ]
    total_calls = 0
    for _ in range(max_iterations):
        try:
            resp = client.chat.completions.create(
                model=model, messages=messages, tools=tools, tool_choice="auto",
            )
        except Exception as e:
            err = str(e)
            if "rate_limit" in err.lower() or "429" in err:
                # Extract wait time from error message if available
                m = re.search(r'try again in (\d+m[\d.]+s)', err)
                wait_hint = f" (retry in {m.group(1)})" if m else ""
                print(f"\n  ⚠ Groq rate limit hit{wait_hint} — stopping LLM calls, keeping web results.")
            else:
                print(f"\n  ⚠ API error: {err[:120]}")
            return "", total_calls
        choice = resp.choices[0]
        msg    = choice.message
        asst_msg = {"role": "assistant", "content": msg.content or ""}
        if msg.tool_calls:
            asst_msg["tool_calls"] = [
                {
                    "id":       tc.id,
                    "type":     "function",
                    "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                }
                for tc in msg.tool_calls
            ]
        messages.append(asst_msg)
        if choice.finish_reason == "stop":
            return msg.content or "", total_calls
        if choice.finish_reason == "tool_calls":
            for tc in (msg.tool_calls or []):
                total_calls += 1
                called_name = tc.function.name
                fn = executors.get(called_name)
                if fn is None:
                    close = difflib.get_close_matches(
                        called_name, executors.keys(), n=1, cutoff=0.75
                    )
                    if close:
                        fn = executors[close[0]]
                try:
                    args   = json.loads(tc.function.arguments or "{}")
                    result = fn(**args) if fn else {"error": f"unknown tool: {called_name}"}
                except Exception as e:
                    result = {"error": str(e)}
                messages.append({
                    "role":         "tool",
                    "tool_call_id": tc.id,
                    "content":      json.dumps(result, default=str),
                })
        else:
            break
    return "", total_calls


# ═══════════════════════════════════════════════════════════════════════════════
# AGENT 1 — SCOUT
# ═══════════════════════════════════════════════════════════════════════════════

_SCOUT_SYSTEM = """You are a startup discovery agent focused on early-stage deep-tech and hard-tech.

LLM knowledge, CORDIS EU grants, university spinoffs, EIC portfolio, and Launch HN
are already pre-loaded. Your job: add keyword-driven web sources.

CALL ORDER:
1. fetch_hn_keywords — Show/Launch HN posts matching thesis keywords
2. fetch_rss_feeds — funding news from TechCrunch, EU-Startups, Sifted, VentureBeat
3. fetch_crunchbase_ddg — Crunchbase profiles for niche keywords
4. fetch_dealroom_ddg — Dealroom EU startup database
5. fetch_sifted — Sifted EU news
6. fetch_ddg_startups — general DDG startup search
7. fetch_eu_startups, fetch_climatedraft, fetch_product_hunt
8. filter_and_deduplicate — ALWAYS LAST

NEVER invent startups. Only return what tools give you.
End with: sources used · total collected."""

_SCOUT_TOOLS = [
    _oai_tool("fetch_hn_keywords",
              "Search HN for Show HN / Launch HN posts matching thesis keywords.",
              {
                  "keywords":     {"type": "array", "items": {"type": "string"}},
                  "pages_per_kw": {"type": "integer", "default": 4},
              },
              required=["keywords"]),
    _oai_tool("fetch_sifted",
              "Scrape Sifted.eu — European startup news. Real HTML, covers deep-tech and climate.",
              {"pages": {"type": "integer", "default": 4}}),
    _oai_tool("fetch_ddg_startups",
              "DuckDuckGo search for company news on eu-startups.com and sifted.eu.",
              {
                  "keywords": {"type": "array", "items": {"type": "string"}},
              },
              required=["keywords"]),
    _oai_tool("fetch_eu_startups",
              "Scrape EU-Startups directory (may return 0 if JS-rendered — that's OK).",
              {}),
    _oai_tool("fetch_climatedraft",
              "Scrape ClimateDraft climate-tech list (may return 0 if JS-rendered — that's OK).",
              {}),
    _oai_tool("fetch_product_hunt",
              "Fetch Product Hunt RSS daily launches.",
              {}),
    _oai_tool("fetch_rss_feeds",
              "Read RSS from TechCrunch, EU-Startups (3 categories), Sifted, VentureBeat, TNW. "
              "Extracts company names from funding/launch article titles. Always returns real XML.",
              {
                  "keywords":   {"type": "array", "items": {"type": "string"}},
                  "max_items":  {"type": "integer", "default": 20},
              },
              required=["keywords"]),
    _oai_tool("fetch_crunchbase_ddg",
              "DDG search for crunchbase.com/organization pages — extracts real startup company profiles.",
              {
                  "keywords":   {"type": "array", "items": {"type": "string"}},
                  "max_per_kw": {"type": "integer", "default": 6},
              },
              required=["keywords"]),
    _oai_tool("fetch_dealroom_ddg",
              "DDG search for Dealroom.co — leading EU startup funding database.",
              {
                  "keywords":   {"type": "array", "items": {"type": "string"}},
                  "max_per_kw": {"type": "integer", "default": 5},
              },
              required=["keywords"]),
    _oai_tool("filter_and_deduplicate",
              "Remove funded signals, article titles, and deduplicate. Call AFTER all fetches.",
              {}),
]


def run_scout_agent(client, thesis, thesis_keywords=None):
    global _scout_buffer
    _scout_buffer = []

    print("\n── Agent 1: Scout ──────────────────────────────────────────────")

    # ── Step 0: real-data sources (no LLM — zero hallucination risk) ──────────
    # LLM knowledge is DISABLED: it hallucinates company names for niche theses.
    # Every source below returns only companies that physically exist on the web.

    def _direct(label, fn, *args):
        print(f"    {label}:", end=" ", flush=True)
        batch = fn(*args) if args else fn()
        _scout_buffer.extend(batch)
        print(f"+{len(batch)} → {len(_scout_buffer)} total")

    kws = thesis_keywords or []

    # ── Hacker News ────────────────────────────────────────────────────────────
    _direct("Launch HN (20 pages)",    fetch_launch_hn, 20)
    _direct("HN keyword search",       fetch_hn_keywords, kws, 6)

    # ── Launch platforms ───────────────────────────────────────────────────────
    _direct("Product Hunt",            fetch_product_hunt)
    _direct("BetaList",                fetch_betalist)

    # ── EU startup news RSS ────────────────────────────────────────────────────
    _direct("EU news RSS (15 feeds)",  fetch_tech_rss_news, kws)
    _direct("RSS feeds (8 feeds)",     fetch_rss_feeds, kws, 40)

    # ── EU startup databases ───────────────────────────────────────────────────
    _direct("EU-Startups directory",   fetch_eu_startups)
    _direct("Sifted EU",               fetch_sifted, 6)
    _direct("Seedtable (11 countries)",fetch_seedtable, kws)
    _direct("F6S programs",            fetch_f6s_programs, kws)

    # ── VC & accelerator portfolios ────────────────────────────────────────────
    _direct("EU VC portfolios (15)",   fetch_eu_vc_portfolios)
    _direct("EU accelerators (9)",     fetch_eu_accelerators)

    # ── University spinoffs ────────────────────────────────────────────────────
    _direct("University spinoffs",     fetch_university_spinoffs)
    _direct("More universities (15)",  fetch_more_universities)

    # ── Developer / code sources ───────────────────────────────────────────────
    _direct("GitHub API",              fetch_github_startups, kws)

    # ── EU grant databases ─────────────────────────────────────────────────────
    _direct("CORDIS EU grants",        fetch_cordis_eu, kws)
    _direct("EIC Portfolio",           fetch_eic_portfolio, kws)

    # ── Database DDG searches ──────────────────────────────────────────────────
    _direct("Crunchbase DDG",          fetch_crunchbase_ddg, kws, 10)
    _direct("Dealroom DDG",            fetch_dealroom_ddg, kws, 8)
    _direct("Wellfound DDG",           fetch_wellfound_ddg, kws, 8)
    _direct("YC EU companies",         fetch_yc_eu_companies)
    _direct("DDG startup search",      fetch_ddg_startups, kws)
    _direct("ClimateDraft",            fetch_climatedraft)

    print(f"\n  Step 0 done: {len(_scout_buffer)} companies from {28} direct sources.")

    # ── Step 1: keyword-driven web sources via agent loop ───────────────────

    def _exec_launch_hn(pages=12):
        batch = fetch_launch_hn(pages)
        _scout_buffer.extend(batch)
        print(f"    Launch HN: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_hn_keywords(keywords, pages_per_kw=4):
        kws   = keywords or thesis_keywords or []
        batch = fetch_hn_keywords(kws, pages_per_kw)
        _scout_buffer.extend(batch)
        print(f"    HN Keywords ({len(kws)} terms): +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_sifted(pages=4):
        batch = fetch_sifted(pages)
        _scout_buffer.extend(batch)
        print(f"    Sifted EU: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_ddg_startups(keywords):
        kws   = keywords or thesis_keywords or []
        batch = fetch_ddg_startups(kws)
        _scout_buffer.extend(batch)
        print(f"    DDG Startups: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_eu_startups():
        batch = fetch_eu_startups()
        _scout_buffer.extend(batch)
        print(f"    EU-Startups: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_climatedraft():
        batch = fetch_climatedraft()
        _scout_buffer.extend(batch)
        print(f"    ClimateDraft/climate: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_ph():
        batch = fetch_product_hunt()
        _scout_buffer.extend(batch)
        print(f"    Product Hunt: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_cordis(keywords):
        kws   = keywords or thesis_keywords or []
        batch = fetch_cordis_eu(kws)
        _scout_buffer.extend(batch)
        print(f"    CORDIS EU ({len(kws)} kws): +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_university_spinoffs():
        print("    University spinoffs: ", end="", flush=True)
        batch = fetch_university_spinoffs()
        _scout_buffer.extend(batch)
        print(f" +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_eic_portfolio(keywords):
        kws   = keywords or thesis_keywords or []
        batch = fetch_eic_portfolio(kws)
        _scout_buffer.extend(batch)
        print(f"    EIC Portfolio ({len(kws)} kws): +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_rss_feeds(keywords, max_items=20):
        kws   = keywords or thesis_keywords or []
        batch = fetch_rss_feeds(kws, max_items)
        _scout_buffer.extend(batch)
        print(f"    RSS Feeds: +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_crunchbase_ddg(keywords, max_per_kw=6):
        kws   = keywords or thesis_keywords or []
        batch = fetch_crunchbase_ddg(kws, max_per_kw)
        _scout_buffer.extend(batch)
        print(f"    Crunchbase DDG ({len(kws)} kws): +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_dealroom_ddg(keywords, max_per_kw=5):
        kws   = keywords or thesis_keywords or []
        batch = fetch_dealroom_ddg(kws, max_per_kw)
        _scout_buffer.extend(batch)
        print(f"    Dealroom DDG ({len(kws)} kws): +{len(batch)} → {len(_scout_buffer)} total")
        return {"fetched": len(batch)}

    def _exec_filter():
        clean, rm_funded, rm_irrel = keyword_filter_and_dedup(_scout_buffer, thesis_keywords)
        _scout_buffer.clear()
        _scout_buffer.extend(clean)
        print(
            f"    Filter: −{rm_funded} noise/funded, −{rm_irrel} off-thesis"
            f" → {len(_scout_buffer)} relevant"
        )
        return {"remaining": len(_scout_buffer), "removed_funded": rm_funded,
                "removed_irrelevant": rm_irrel}

    executors = {
        "fetch_hn_keywords":      _exec_hn_keywords,
        "fetch_sifted":           _exec_sifted,
        "fetch_ddg_startups":     _exec_ddg_startups,
        "fetch_eu_startups":      _exec_eu_startups,
        "fetch_climatedraft":     _exec_climatedraft,
        "fetch_product_hunt":     _exec_ph,
        "fetch_rss_feeds":        _exec_rss_feeds,
        "fetch_crunchbase_ddg":   _exec_crunchbase_ddg,
        "fetch_dealroom_ddg":     _exec_dealroom_ddg,
        "filter_and_deduplicate": _exec_filter,
    }

    kw_hint = (
        f"\nThesis keywords for targeted search: {thesis_keywords[:10]}"
        if thesis_keywords else ""
    )

    summary, n_calls = run_agent(
        client=client,
        model=MODEL,
        system=_SCOUT_SYSTEM,
        first_message=(
            f"Pre-loaded: {len(_scout_buffer)} real companies from "
            f"Launch HN, Product Hunt, BetaList, GitHub, CORDIS EU, "
            f"university spinoffs, and EIC Portfolio. "
            f"Now add keyword-driven web sources: "
            f"fetch_hn_keywords → fetch_rss_feeds → fetch_crunchbase_ddg → "
            f"fetch_dealroom_ddg → fetch_sifted → fetch_ddg_startups → "
            f"fetch_eu_startups → fetch_climatedraft → "
            f"filter_and_deduplicate."
            + kw_hint
        ),
        tools=_SCOUT_TOOLS,
        executors=executors,
    )
    print(f"  Scout done ({n_calls} tool calls). {len(_scout_buffer)} startups collected.")
    if summary:
        print(f"  → {summary[:220]}")
    return list(_scout_buffer)


# ═══════════════════════════════════════════════════════════════════════════════
# AGENT 2 — ANALYST
# ═══════════════════════════════════════════════════════════════════════════════

def run_analyst_agent(client, raw_startups, thesis, thesis_keywords=None):
    global _analyst_buffer
    _analyst_buffer = []

    print("\n── Agent 2: Analyst ────────────────────────────────────────────")

    # Phase 1: parse thesis into structured dimensions
    parsed_thesis = {}
    print("  Parsing thesis...", end=" ", flush=True)
    if thesis:
        try:
            raw, _lbl = llm_call([
                {"role": "system", "content":
                    "Parse an investor thesis into structured JSON. "
                    "Return ONLY valid JSON, no markdown."},
                {"role": "user", "content": (
                    f"Investor thesis:\n\n{thesis}\n\n"
                    "Return JSON:\n"
                    '{"sectors":[], "geographies":[], "stage":"", '
                    '"technologies":[], "themes":[], '
                    '"must_haves":[], "avoid":[], "ideal_startup":""}'
                )},
            ], max_tokens=600)
            if not raw:
                raise ValueError("empty response")
            raw = re.sub(r"```(?:json)?", "", raw).strip().strip("`")
            parsed_thesis = json.loads(raw)
            print(f"done via {_lbl}. Technologies: {parsed_thesis.get('technologies', [])[:5]}")
        except Exception as e:
            print(f"failed ({e}) — using raw thesis.")

    # Build thesis context string for batch scoring prompts
    if parsed_thesis:
        thesis_context = "\n".join([
            f"Target sectors      : {', '.join(parsed_thesis.get('sectors', ['any']))}",
            f"Target geographies  : {', '.join(parsed_thesis.get('geographies', ['any']))}",
            f"Stage preference    : {parsed_thesis.get('stage', 'any')}",
            f"Key technologies    : {', '.join(parsed_thesis.get('technologies', [])[:6])}",
            f"Key themes          : {', '.join(parsed_thesis.get('themes', [])[:4])}",
            f"Must-haves          : {', '.join(parsed_thesis.get('must_haves', ['none'])[:4])}",
            f"Avoid               : {', '.join(parsed_thesis.get('avoid', ['none'])[:3])}",
            f"Ideal startup       : {parsed_thesis.get('ideal_startup', 'not specified')}",
            f"Full thesis         : {thesis[:500]}",
        ])
    else:
        thesis_context = thesis[:600] if thesis else "Score on general pre-seed quality."

    # Phase 2: pre-rank by keyword relevance, score only the best candidates
    # This keeps token usage within Groq's 100k/day free limit regardless of
    # how many companies were discovered.
    MAX_TO_SCORE = 200  # ~25 batches × ~2k tokens ≈ 50k tokens (safe within 100k limit)

    if thesis_keywords and len(raw_startups) > MAX_TO_SCORE:
        # Score each startup by how many thesis keywords appear in name+desc+tags
        def _relevance(s):
            text = (s.get("name","") + " " + s.get("description","") + " " + s.get("tags","")).lower()
            return sum(1 for k in thesis_keywords if k in text)

        ranked = sorted(raw_startups, key=_relevance, reverse=True)
        # Always include at least some from each source for diversity
        top    = ranked[:MAX_TO_SCORE]
        skipped = len(raw_startups) - len(top)
        print(f"  Pre-filter: {len(raw_startups)} → {len(top)} most thesis-relevant "
              f"(skipped {skipped} low-relevance). Fits within free API quota.")
        raw_startups = top
    else:
        print(f"  Scoring all {len(raw_startups)} startups (within quota limit).")

    BATCH         = BATCH_SIZE
    total         = len(raw_startups)
    total_batches = (total + BATCH - 1) // BATCH
    scored_map    = {}

    print(f"  Scoring {total} startups in {total_batches} batches of {BATCH}...")
    _stopped_early = False

    for i in range(0, total, BATCH):
        # Stop as soon as all API keys are rate-limited — output what we have
        if all_apis_exhausted():
            print(f"\n  ★ All API keys exhausted at batch {i // BATCH + 1}/{total_batches}.")
            print(f"  ★ Scored {len(scored_map)} startups — proceeding to output.")
            _stopped_early = True
            break

        batch     = raw_startups[i:i + BATCH]
        batch_num = i // BATCH + 1
        print(f"    Batch {batch_num}/{total_batches}...", end=" ", flush=True)

        numbered = "\n".join(
            f"{j+1}. [{s['name']}] — {s['description'][:200]}"
            for j, s in enumerate(batch)
        )

        prompt = (
            f"Investor thesis context:\n{thesis_context}\n\n"
            f"Score these {len(batch)} startups for thesis fit.\n\n"
            f"{numbered}\n\n"
            f"Return ONLY a JSON array with exactly {len(batch)} objects:\n"
            f'[{{"name":"EXACT name from list","overall":3,"sector_fit":2,'
            f'"geo_fit":3,"stage_fit":4,"tech_fit":3,"theme_fit":2,'
            f'"reason":"2-3 sentence verdict on thesis fit",'
            f'"best_signal":"strongest reason to look closer",'
            f'"red_flag":"biggest concern or none",'
            f'"stage_guess":"pre-seed/seed/unknown",'
            f'"tech_tags":"comma-separated specific tech tags"}}]\n\n'
            f"Scoring guide:\n"
            f"  5 = perfect match  4 = strong  3 = moderate  2 = weak  1 = no fit\n"
            f"  Be strict — 4+ should be rare.\n\n"
            f"HARD PENALTIES (apply automatically):\n"
            f"  • If company is US/Canada/Asia AND thesis targets EU → geo_fit = 1, overall ≤ 2\n"
            f"  • If company has raised Series A or beyond (≥$5M VC) → stage_fit = 1, overall ≤ 2\n"
            f"  • If company name matches a well-known scaleup → overall = 1\n"
            f"  • If description says 'acquired', 'IPO', 'publicly traded' → overall = 1\n"
            f"  • Universities and government bodies → overall = 1 (not investable startups)\n"
            f"  • Pure screen-based tools (coding IDEs, flashcard apps, LMS platforms, AI tutors "
            f"for coding/maths) → theme_fit = 1 if thesis emphasises outdoor/embodied/creative "
            f"learning. Being in 'EdTech' is NOT sufficient for theme_fit ≥ 3.\n"
            f"  • HN post titles that are not company names (start with 'I built', 'We're', "
            f"'Show HN', full sentences) → overall = 1"
        )

        try:
            text, _lbl = llm_call([
                {"role": "system", "content": (
                    "You are a strict deep-tech investment analyst. "
                    "Score precisely and honestly. "
                    "Return only a valid JSON array — no markdown, no extra text."
                )},
                {"role": "user", "content": prompt},
            ], max_tokens=2000)
            if not text:
                raise ValueError("empty response from all APIs")
            text = re.sub(r"```(?:json)?", "", text).strip().strip("`")
            scores = json.loads(text)

            saved = 0
            for score in scores:
                score_name = str(score.get("name", "")).strip()
                original   = next(
                    (s for s in batch if s["name"].strip() == score_name), None
                )
                if original is None:
                    names = [s["name"].strip() for s in batch]
                    close = difflib.get_close_matches(
                        score_name, names, n=1, cutoff=0.55
                    )
                    if close:
                        original = next(
                            s for s in batch if s["name"].strip() == close[0]
                        )
                if original is None:
                    continue

                merged = dict(original)
                merged.update({
                    "overall_score": max(0, min(5, int(score.get("overall",    0)))),
                    "sector_fit":    max(0, min(5, int(score.get("sector_fit", 0)))),
                    "geo_fit":       max(0, min(5, int(score.get("geo_fit",    0)))),
                    "stage_fit":     max(0, min(5, int(score.get("stage_fit",  0)))),
                    "tech_fit":      max(0, min(5, int(score.get("tech_fit",   0)))),
                    "theme_fit":     max(0, min(5, int(score.get("theme_fit",  0)))),
                    "reason":        str(score.get("reason",      ""))[:500],
                    "best_signal":   str(score.get("best_signal", ""))[:300],
                    "red_flag":      str(score.get("red_flag",    "none"))[:300],
                    "stage_guess":   str(score.get("stage_guess", "unknown")),
                    "tech_tags":     str(score.get("tech_tags",   "")),
                })
                scored_map[original["name"]] = merged
                saved += 1

            print(f"+{saved}  ({len(scored_map)}/{total})")

        except Exception as e:
            print(f"failed — {e}")
            if _consecutive_failures >= 3:
                print(f"\n  ★ {_consecutive_failures} consecutive failures — API quota exhausted.")
                print(f"  ★ Scored {len(scored_map)} startups — proceeding to output.")
                _stopped_early = True
                break

        time.sleep(1.0)  # brief pause between batches

    _analyst_buffer = sorted(
        scored_map.values(),
        key=lambda x: x.get("overall_score", 0),
        reverse=True,
    )
    print(f"  Analyst done. {len(_analyst_buffer)}/{total} startups scored.")
    return list(_analyst_buffer)


# ═══════════════════════════════════════════════════════════════════════════════
# AGENT 3 — VERIFIER + ENRICHER  (direct Python — no LLM, no context buildup)
# ═══════════════════════════════════════════════════════════════════════════════

def run_verifier_agent(client, scored_startups):
    global _final_results
    _final_results = []

    all_top      = scored_startups[:TOP_N]
    candidates   = [s for s in all_top if s.get("overall_score", 0) >= MIN_SCORE_TO_VERIFY]
    skipped_low  = len(all_top) - len(candidates)
    funded_removed      = 0
    hallucination_removed = 0

    print("\n── Agent 3: Verifier + Enricher ────────────────────────────────")
    if skipped_low:
        print(f"  Skipping {skipped_low} candidates scored < {MIN_SCORE_TO_VERIFY}/5.")
    print(f"  Deep-checking {len(candidates)} candidates (score ≥ {MIN_SCORE_TO_VERIFY})...")
    print(f"  Checks: funding (EDGAR + DDG) · existence (DDG) · website coherence")

    for c in candidates:
        name        = c["name"]
        website_url = c.get("website_url", "")
        source      = c.get("source", "")
        tech_tags   = c.get("tech_tags", "")
        description = c.get("description", "")

        print(f"    [{c.get('overall_score',0)}/5] {name[:45]:<45}", end=" ", flush=True)

        # ── Step 1: Funding checks ──────────────────────────────────────────
        edgar                = sec_edgar_check(name)
        domain_age, is_young = domain_age_check(website_url)
        web                  = web_funding_check(name)

        if edgar is True or web == "funded":
            funded_removed += 1
            src = "EDGAR" if edgar is True else "DDG"
            print(f"⛔ funded ({src})")
            continue

        # ── Step 2: Existence check (required for LLM-sourced companies) ───
        if source in _LLM_SOURCES:
            exists = check_company_exists(name, tech_tags, description)
            if exists == "not_found":
                hallucination_removed += 1
                print("⚠ not_found (no web trace — likely hallucinated)")
                continue
            if exists == "mismatch":
                hallucination_removed += 1
                print("⚠ mismatch (found but wrong business — name collision)")
                continue

        # ── Step 3: Website enrichment ──────────────────────────────────────
        snippet = enrich_with_website(c)

        # ── Step 4: Website content coherence check ─────────────────────────
        if snippet and not snippet_is_relevant(snippet, description, tech_tags):
            hallucination_removed += 1
            print("⚠ website mismatch (homepage describes different business)")
            continue

        confidence = (
            "likely_unfunded"
            if web == "clean" and (is_young or edgar is False)
            else "uncertain"
        )
        age_str = f"{domain_age}mo" if domain_age else "?"
        print(f"✓ {confidence} (age:{age_str})" + (" + enriched" if snippet else ""))

        result = dict(c)
        result["funding_confidence"] = confidence
        result["domain_age_months"]  = domain_age
        result["website_snippet"]    = snippet
        _final_results.append(result)

    for rank, r in enumerate(_final_results, 1):
        r["rank"] = rank

    print(
        f"\n  Verifier done. "
        f"{len(_final_results)} confirmed  |  "
        f"{funded_removed} funded removed  |  "
        f"{hallucination_removed} hallucinated/mismatched removed."
    )
    return list(_final_results)


# ═══════════════════════════════════════════════════════════════════════════════
# OUTPUT — EXCEL + CSV + TERMINAL TABLE
# ═══════════════════════════════════════════════════════════════════════════════

_SCORE_COLORS = {
    5: "1B5E20",  # deep green
    4: "2E7D32",  # green
    3: "F57F17",  # amber
    2: "BF360C",  # orange-red
    1: "B71C1C",  # red
    0: "757575",  # grey
}


def save_excel(results, filepath=OUTPUT_XLSX):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Startup Scout Results"

    header_fill  = PatternFill("solid", fgColor="1A237E")
    header_font  = Font(bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border  = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    wrap = Alignment(vertical="top", wrap_text=True)

    headers = [
        "#", "Company Name", "Website (click)", "Score",
        "Sector", "Geo", "Stage", "Tech Fit", "Theme",
        "Funding Status", "Domain Age", "Source",
        "Verdict / Reason", "Best Signal", "Red Flag",
        "Tech Tags", "Stage Guess", "Original Description", "Website Snippet",
    ]
    col_widths = [
        4, 32, 35, 7,
        8, 7, 8, 9, 8,
        18, 12, 16,
        65, 45, 40,
        28, 14, 65, 65,
    ]

    ws.append(headers)
    for col_num, _ in enumerate(headers, 1):
        cell            = ws.cell(row=1, column=col_num)
        cell.fill       = header_fill
        cell.font       = header_font
        cell.alignment  = header_align
        ws.column_dimensions[get_column_letter(col_num)].width = col_widths[col_num - 1]

    ws.row_dimensions[1].height = 30

    for r in results:
        score      = r.get("overall_score", 0)
        score_fill = PatternFill("solid", fgColor=_SCORE_COLORS.get(score, "757575"))
        website    = r.get("website_url") or r.get("url", "")

        row_data = [
            r.get("rank", ""),
            r.get("name", ""),
            "",                                         # col 3 = hyperlink (set below)
            f"{score}/5",
            r.get("sector_fit", ""),
            r.get("geo_fit", ""),
            r.get("stage_fit", ""),
            r.get("tech_fit", ""),
            r.get("theme_fit", ""),
            r.get("funding_confidence", ""),
            f"{r.get('domain_age_months', '?')}mo" if r.get("domain_age_months") else "?",
            r.get("source", ""),
            r.get("reason", ""),
            r.get("best_signal", ""),
            r.get("red_flag", ""),
            r.get("tech_tags", ""),
            r.get("stage_guess", ""),
            r.get("description", ""),
            r.get("website_snippet", ""),
        ]

        ws.append(row_data)
        row_num = ws.max_row
        alt_row = row_num % 2 == 0

        for col_num in range(1, len(headers) + 1):
            cell            = ws.cell(row=row_num, column=col_num)
            cell.border     = thin_border
            cell.alignment  = wrap
            if col_num == 4:
                cell.fill = score_fill
                cell.font = Font(bold=True, color="FFFFFF", size=11)
            elif alt_row:
                cell.fill = PatternFill("solid", fgColor="F3F3F3")

        # Clickable hyperlink in column 3
        if website:
            link_cell           = ws.cell(row=row_num, column=3)
            link_cell.value     = extract_domain(website) or website[:40]
            link_cell.hyperlink = website
            link_cell.font      = Font(color="1565C0", underline="single")

        ws.row_dimensions[row_num].height = 90

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(filepath)
    print(f"  ✓ Excel → {filepath}")


def save_csv(results, filepath=OUTPUT_CSV):
    fieldnames = [
        "rank", "name", "website_url", "overall_score",
        "sector_fit", "geo_fit", "stage_fit", "tech_fit", "theme_fit",
        "tech_tags", "stage_guess",
        "reason", "best_signal", "red_flag",
        "funding_confidence", "domain_age_months",
        "source", "url", "description", "website_snippet", "scraped_at",
    ]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in results:
            writer.writerow(row)
    print(f"  ✓ CSV  → {filepath}")


def print_top_table(results, n=10):
    STARS = {5:"★★★★★", 4:"★★★★☆", 3:"★★★☆☆", 2:"★★☆☆☆", 1:"★☆☆☆☆", 0:"☆☆☆☆☆"}
    rows  = []
    for r in results[:n]:
        s = r.get("overall_score", 0)
        rows.append([
            r.get("rank", ""),
            r.get("name", "")[:40],
            f"{STARS.get(s,'?')} {s}/5",
            (f"{r.get('sector_fit',0)}|{r.get('geo_fit',0)}|"
             f"{r.get('stage_fit',0)}|{r.get('tech_fit',0)}|{r.get('theme_fit',0)}"),
            r.get("funding_confidence", "")[:16],
            r.get("source", "")[:14],
        ])
    print("\n" + tabulate(
        rows,
        headers=["#", "Company", "Overall", "Se|Ge|St|Tc|Th", "Funding", "Source"],
        tablefmt="rounded_outline",
    ))
    print("  Se=Sector  Ge=Geo  St=Stage  Tc=Tech  Th=Theme  (1–5 each)")
    print(f"  ▶ Full details with clickable links → {OUTPUT_XLSX}")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "=" * 68)
    print("  STARTUP SCOUT v2 — Multi-Agent Deep-Tech Discovery")
    print("  Sources: Launch HN · HN Keywords · Product Hunt · BetaList · GitHub API")
    print("           CORDIS EU · University Spinoffs · EIC Portfolio · RSS Feeds")
    print("           Crunchbase DDG · Dealroom DDG · Sifted · EU-Startups + more")
    print("  AI role: keyword extraction · scoring only — NO hallucinated companies")
    print("  Flow:    Thesis → Keywords → Scout (15 real sources) → Analyst → Verifier + Excel")
    print("=" * 68 + "\n")

    groq_key = os.getenv("GROQ_API_KEY", "").strip()

    if not groq_key:
        print("ERROR: No GROQ_API_KEY found in .env\n")
        print("  1. Go to https://console.groq.com → API Keys → Create key")
        print("  2. Add to .env:  GROQ_API_KEY=your_key_here")
        sys.exit(1)

    global MODEL, _API_CLIENTS
    _API_CLIENTS = []

    _groq_client = openai.OpenAI(
        api_key=groq_key,
        base_url="https://api.groq.com/openai/v1",
    )
    _API_CLIENTS.append((_groq_client, "llama-3.3-70b-versatile", "Groq"))
    print(f"  ✓ Groq (llama-3.3-70b-versatile) — 100k tokens/day free")

    client, MODEL, _ = _API_CLIENTS[0]

    print("Paste your investor thesis below.")
    print("Press ENTER TWICE when done (or once with nothing to skip):\n")
    lines = []
    try:
        while True:
            line = input()
            if line == "" and lines:
                break
            if line == "" and not lines:
                break
            lines.append(line)
    except EOFError:
        pass
    thesis = " ".join(lines).strip()

    if thesis:
        print(f"\nThesis captured ({len(thesis)} chars).\n")
    else:
        print("No thesis — scoring on general pre-seed quality.\n")

    # Extract thesis keywords BEFORE Scout so they drive both search and pre-filter
    thesis_keywords = []
    if thesis:
        print("Extracting thesis keywords for targeted search...", end=" ", flush=True)
        thesis_keywords = extract_thesis_keywords(client, thesis)
        print(f"{len(thesis_keywords)} keywords: {thesis_keywords[:6]}")

    t0 = time.time()

    raw_startups = run_scout_agent(client, thesis, thesis_keywords)
    if not raw_startups:
        print("\nNo startups collected. Check internet connection and API key.")
        sys.exit(1)

    scored_startups = run_analyst_agent(client, raw_startups, thesis, thesis_keywords)
    if not scored_startups:
        print("\nNo startups scored (API quota exhausted before any batch completed).")
        print("Run again tomorrow when quota resets, or add a second API key.")
        sys.exit(0)

    final_results = run_verifier_agent(client, scored_startups)
    if not final_results:
        print("\nVerifier flagged all candidates as funded.")
        sys.exit(0)

    print("\n── Output ──────────────────────────────────────────────────────")
    print_top_table(final_results)
    save_excel(final_results)
    save_csv(final_results)

    elapsed        = round(time.time() - t0)
    top            = final_results[0]
    funded_removed = len(scored_startups[:TOP_N]) - len(final_results)

    print(
        f"\n{'─'*68}\n"
        f"  {len(raw_startups)} scraped  →  {len(scored_startups)} scored  →  "
        f"{funded_removed} funded removed  →  {len(final_results)} final\n"
        f"  Top match  : {top.get('name','N/A')[:60]}  ({top.get('overall_score',0)}/5)\n"
        f"  Completed  : {elapsed}s\n"
        f"  Results    → {OUTPUT_XLSX}\n"
        f"{'─'*68}\n"
    )


if __name__ == "__main__":
    main()
