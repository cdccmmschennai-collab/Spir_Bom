"""
Resolves a manufacturer's home country (name + short code, per Reference.xlsx
in the project root) for the "Manufacturer COUNTRY NAME" / "Manufacturer
Country CODE" columns in the SAP OUTPUT file.

The SPIR file itself never states a manufacturer's country (only the
manufacturer's *name*, e.g. "ABB LLC"), so this looks the country up two
ways, in order:
  1. Wikidata's own entity search (wbsearchentities) -- fast, precise when it
     works, but its relevance ranking is sometimes wrong for a short company
     name (e.g. 'ABB' ranks the band 'ABBA' first, 'MOXA' ranks the drug
     'amoxicillin' first) and can miss smaller manufacturers entirely.
  2. A Wikipedia full-text search (official public search API, no scraping,
     no API key) for the manufacturer name -- this is effectively the
     "search the web for the manufacturer" step: Wikipedia's search finds
     companies Wikidata's own search missed (e.g. 'ATEN International',
     'Moxa Technologies'), and each hit is resolved back to its Wikidata
     item to reuse the exact same country-claim logic as tier 1.
Both tiers only accept a candidate whose description/snippet reads like an
actual company, and cross-reference the resulting country name against
Reference.xlsx's Country code sheet (via engine.reference_data) to get the
short code. Results are cached on disk (data/manufacturer_country_cache.json)
so a manufacturer is only ever looked up once. Anything that can't be
resolved with confidence by either tier is left blank and appended to
data/manufacturer_country_review.log for manual follow-up -- this never
guesses.

The vendor contact block is deliberately NOT used as a country source: it's
the LOCAL vendor's address, not the manufacturer's actual country (e.g. ABB's
Doha office vs. ABB's home country, Switzerland).
"""
import os
import re
import json
import time
import datetime
import urllib.request
import urllib.parse
import urllib.error

from . import reference_data

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT_DIR, 'data')
CACHE_PATH = os.path.join(DATA_DIR, 'manufacturer_country_cache.json')
REVIEW_LOG_PATH = os.path.join(DATA_DIR, 'manufacturer_country_review.log')
CORRECTIONS_PATH = os.path.join(DATA_DIR, 'manufacturer_name_corrections.json')

# Data-entry typos in SPIR manufacturer fields that have been manually
# confirmed (not guessed): the misspelling never resolves on its own, but
# the corrected name goes through the full web-search + Reference Excel
# pipeline exactly like any other manufacturer -- this only fixes the name,
# never the country. Add more entries here, or in
# data/manufacturer_name_corrections.json (which is merged in and lets you
# add confirmed corrections without a code change), as you confirm them.
_DEFAULT_NAME_CORRECTIONS = {
    'CIRCORE': 'CIRCOR',
}
_name_corrections = None    

_SUFFIX_RE = re.compile(
    r'[,.]?\s*\b(LLC|L\.L\.C|INC|INCORPORATED|LTD|LIMITED|LLP|CO|CORP|CORPORATION|'
    r'GMBH|AG|S\.?A|N\.?V|BV|PVT|PLC|SPA|KG|OY|AB|AS|SRL)\.?\s*$',
    re.IGNORECASE)

_cache = None             # normalized manufacturer name -> [country_name, code], loaded once from disk
_failed_this_run = set()  # normalized names that failed already, so one run doesn't
                          # re-hit the network for every spare row of the same manufacturer


def _load_cache():
    global _cache
    if _cache is not None:
        return _cache
    _cache = {}
    if os.path.isfile(CACHE_PATH):
        try:
            with open(CACHE_PATH, 'r', encoding='utf-8') as f:
                _cache = json.load(f)
        except Exception:
            _cache = {}
    return _cache


def _save_cache():
    os.makedirs(DATA_DIR, exist_ok=True)
    try:
        with open(CACHE_PATH, 'w', encoding='utf-8') as f:
            json.dump(_cache, f, indent=2, sort_keys=True)
    except Exception:
        pass


def _load_name_corrections():
    global _name_corrections
    if _name_corrections is not None:
        return _name_corrections
    m = dict(_DEFAULT_NAME_CORRECTIONS)
    if os.path.isfile(CORRECTIONS_PATH):
        try:
            with open(CORRECTIONS_PATH, 'r', encoding='utf-8') as f:
                m.update({k.upper(): v for k, v in json.load(f).items()})
        except Exception:
            pass
    _name_corrections = m
    return m


def _clean_name(name: str) -> str:
    """'ABB LLC' -> 'ABB' -- strips common corporate suffixes so the search
    query targets the actual company name, then applies any confirmed
    data-entry-typo correction (see _DEFAULT_NAME_CORRECTIONS above)."""
    name = str(name or '').strip()
    prev = None
    while prev != name:
        prev = name
        name = _SUFFIX_RE.sub('', name).strip().rstrip(',').strip()
    return _load_name_corrections().get(name.upper(), name)


def _log_for_review(manufacturer: str, reason: str):
    os.makedirs(DATA_DIR, exist_ok=True)
    try:
        with open(REVIEW_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f'{datetime.datetime.utcnow().isoformat()}\t{manufacturer}\t{reason}\n')
    except Exception:
        pass


def _api_get(base_url: str, params: dict):
    url = base_url + '?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent': 'bom-tool/1.0'})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=6) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            if e.code in (429, 503) and attempt < 2:
                time.sleep(2 * (attempt + 1))
                continue
            raise


def _wikidata_get(params: dict):
    return _api_get('https://www.wikidata.org/w/api.php', params)


def _wikipedia_get(params: dict):
    return _api_get('https://en.wikipedia.org/w/api.php', params)


# Reference.xlsx spells most countries normally, but uses a handful of old or
# abbreviated names that no longer match Wikidata's current English label for
# the same country (e.g. it predates the 2019/2018 Macedonia/Swaziland
# renames, and shortens a few others). This is a static reconciliation table
# between two fixed country lists -- not a per-manufacturer guess.
_COUNTRY_NAME_ALIASES = {
    'united states': 'usa',
    'russia': 'russian fed.',
    'united arab emirates': 'utd.arab emir.',
    'czechia': 'czech republic',
    'north macedonia': 'macedonia',
    'eswatini': 'swaziland',
    'belarus': 'white russia',
    'moldova': 'moldavia',
    'brunei': 'brunei dar-es-s',
    'bosnia and herzegovina': 'bosnia-herz.',
    'dominican republic': 'dominican rep.',
    "côte d'ivoire": 'ivory coast',
    'federated states of micronesia': 'micronesia',
}

_BUSINESS_KEYWORDS = ('company', 'corporation', 'manufacturer', 'manufactures', 'conglomerate',
                      'multinational', 'business', 'enterprise', 'firm', 'brand', 'group',
                      'industrial', 'industries', 'holding', 'producer', 'supplier')


def _pick_candidate(candidates, text_of):
    """A plain top search hit is often wrong for a short company name (e.g.
    searching 'ABB' ranks the band 'ABBA' first, 'MOXA' ranks the drug
    'amoxicillin'). Only accept a candidate whose description/snippet (via
    `text_of`) reads like an actual company -- if none of the top hits look
    like a company, this returns None rather than guessing from an unrelated
    entity."""
    def is_business(c):
        return any(k in (text_of(c) or '').lower() for k in _BUSINESS_KEYWORDS)

    business = [c for c in candidates if is_business(c)]
    return business[0] if business else None


def _first_claim_target(claims_resp: dict):
    for statements in (claims_resp.get('claims') or {}).values():
        for st in statements:
            try:
                return st['mainsnak']['datavalue']['value']['id']
            except (KeyError, TypeError):
                continue
    return None


def _country_name_for_qid(qid: str):
    """Given a Wikidata entity QID, chases P17 (country), falling back to
    P159 (headquarters location) -> that place's own P17, and returns the
    resolved country's English label. None if no claim leads anywhere."""
    claims = _wikidata_get({'action': 'wbgetclaims', 'entity': qid, 'property': 'P17', 'format': 'json'})
    country_qid = _first_claim_target(claims)

    if not country_qid:
        hq_claims = _wikidata_get({'action': 'wbgetclaims', 'entity': qid, 'property': 'P159', 'format': 'json'})
        hq_qid = _first_claim_target(hq_claims)
        if hq_qid:
            hq_country_claims = _wikidata_get({'action': 'wbgetclaims', 'entity': hq_qid, 'property': 'P17', 'format': 'json'})
            country_qid = _first_claim_target(hq_country_claims)

    if not country_qid:
        return None

    labels = _wikidata_get({'action': 'wbgetentities', 'ids': country_qid, 'props': 'labels', 'languages': 'en', 'format': 'json'})
    entity = (labels.get('entities') or {}).get(country_qid) or {}
    return ((entity.get('labels') or {}).get('en') or {}).get('value')


def _lookup_country_via_wikidata(clean_name: str):
    """Tier 1: Wikidata's own entity search. Best-effort country NAME (e.g.
    'Switzerland') for a company name, or None if no confident match was
    found."""
    search = _wikidata_get({
        'action': 'wbsearchentities', 'search': clean_name, 'language': 'en',
        'format': 'json', 'type': 'item', 'limit': 5,
    })
    candidates = search.get('search') or []
    best = _pick_candidate(candidates, lambda c: c.get('description'))
    if not best:
        return None
    return _country_name_for_qid(best['id'])


def _lookup_country_via_wikipedia(clean_name: str):
    """Tier 2 (web-search fallback): searches Wikipedia's public full-text
    index for the manufacturer name -- this finds real companies Wikidata's
    own entity search misses (e.g. 'ATEN International', 'Moxa
    Technologies'), whose Wikipedia pages are then resolved back to their
    linked Wikidata item to reuse the same country-claim logic. Only a
    result whose search snippet reads like an actual company is accepted."""
    # A short/generic manufacturer name (e.g. 'ATEN', 'MOXA') alone often
    # loses to older, more-linked unrelated topics in Wikipedia's ranking;
    # qualifying the query narrows results toward the actual company.
    search = _wikipedia_get({
        'action': 'query', 'list': 'search', 'srsearch': f'{clean_name} company',
        'format': 'json', 'srlimit': 5, 'srprop': 'snippet',
    })
    results = (search.get('query') or {}).get('search') or []

    def snippet_text(r):
        return re.sub(r'<[^>]+>', '', r.get('snippet') or '')

    best = _pick_candidate(results, snippet_text)
    if not best:
        return None

    props = _wikipedia_get({'action': 'query', 'titles': best['title'], 'prop': 'pageprops', 'format': 'json'})
    pages = ((props.get('query') or {}).get('pages') or {})
    qid = next(iter(pages.values()), {}).get('pageprops', {}).get('wikibase_item')
    if not qid:
        return None
    return _country_name_for_qid(qid)


def _resolve_single(clean_name: str):
    """(name, code) for one already-cleaned manufacturer name via the
    cache/network pipeline. Does not cache or log a failure -- the caller
    decides that once every candidate name has been tried (see
    resolve_country, which may try more than one name for a combined
    field like 'PERTOFAC/ SULZER')."""
    key = clean_name.lower()
    if not key:
        return '', ''

    cache = _load_cache()
    if key in cache:
        return tuple(cache[key])
    if key in _failed_this_run:
        return '', ''

    name, code = '', ''
    country_name = None
    try:
        country_name = _lookup_country_via_wikidata(clean_name)
    except Exception:
        pass
    if not country_name:
        try:
            country_name = _lookup_country_via_wikipedia(clean_name)
        except Exception:
            pass

    if country_name:
        lookup_key = country_name.strip().lower()
        lookup_key = _COUNTRY_NAME_ALIASES.get(lookup_key, lookup_key)
        resolved_code = reference_data.country_code(lookup_key)
        if resolved_code:
            name, code = country_name.strip(), resolved_code

    if code:
        cache[key] = [name, code]
        _save_cache()
    else:
        # Not cached to disk on purpose -- a failure (network hiccup, no
        # confident match) should be retried on the next run rather than
        # permanently locked in as blank.
        _failed_this_run.add(key)
    return name, code


def resolve_country(manufacturer_name):
    """Best-effort (country_name, country_code) for a manufacturer's home
    country -- e.g. ('Switzerland', 'CH') for 'ABB'. Returns ('', '') if it
    can't be resolved with confidence -- never guesses, and logs the
    manufacturer to data/manufacturer_country_review.log for manual
    follow-up.

    A SPIR sometimes names more than one manufacturer/brand in one field,
    e.g. 'PERTOFAC/ SULZER' -- each '/'-separated name is tried in turn
    rather than searching the combined, unsearchable string as a whole."""
    raw = str(manufacturer_name or '').strip()
    if not raw:
        return '', ''

    candidates = [_clean_name(p) for p in raw.split('/') if p.strip()] or [_clean_name(raw)]
    for clean in candidates:
        if not clean:
            continue
        name, code = _resolve_single(clean)
        if code:
            return name, code

    _log_for_review(raw, 'country not resolved with confidence')
    return '', ''


def get_country_code(manufacturer_name) -> str:
    """Backward-compatible convenience wrapper: just the short code."""
    return resolve_country(manufacturer_name)[1]
