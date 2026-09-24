"""
Fetches live currency -> QAR conversion rates. Falls back to the official
USD/QAR peg (~3.64) if the live lookup fails, so the tool never hard-crashes
on a network hiccup, but always prefers a fresh rate when it can get one.
"""
import urllib.request
import json

FALLBACK_RATES = {
    'USD': 3.64,   # official long-standing peg
}

_cache = {}


def get_rate_to_qar(currency_code: str) -> float | None:
    code = (currency_code or '').strip().upper()
    if not code:
        return None
    if code == 'QAR':
        return 1.0
    if code in _cache:
        return _cache[code]

    rate = None
    try:
        url = f'https://api.exchangerate-api.com/v4/latest/{code}'
        with urllib.request.urlopen(url, timeout=5) as resp:
            data = json.loads(resp.read().decode())
            rate = data.get('rates', {}).get('QAR')
    except Exception:
        rate = None

    if rate is None:
        rate = FALLBACK_RATES.get(code)

    if rate is not None:
        _cache[code] = rate
    return rate
