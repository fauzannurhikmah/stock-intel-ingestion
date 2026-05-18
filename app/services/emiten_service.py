from __future__ import annotations

from math import ceil
from typing import Any, Dict, List, Optional, Tuple
import cloudscraper
import requests
from config.settings import REQUEST_TIMEOUT


AJAIB_STOCK_LIST_URL = "https://ajaib.co.id/api/stock-list"
IDX_PROFILE_URL = "https://www.idx.co.id/primary/ListedCompany/GetCompanyProfilesDetail"

AJAIB_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://ajaib.co.id/",
}

IDX_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://www.idx.co.id/",
}

_IDX_SCRAPER = None


def _get_json(url: str, params: Optional[dict] = None, headers: Optional[dict] = None) -> dict:
    response = requests.get(url, params=params or {}, headers=headers or AJAIB_HEADERS, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.json()


def _get_idx_scraper(reset: bool = False):
    global _IDX_SCRAPER
    if _IDX_SCRAPER is None or reset:
        _IDX_SCRAPER = cloudscraper.create_scraper(
            browser={"browser": "chrome", "platform": "windows", "mobile": False}
        )
    return _IDX_SCRAPER


def _get_json_cloudscraper(url: str, params: Optional[dict] = None, headers: Optional[dict] = None) -> dict:
    scraper = _get_idx_scraper()
    try:
        response = scraper.get(url, params=params or {}, headers=headers or AJAIB_HEADERS, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        return response.json()
    except Exception:
        # Recreate scraper once; stale or blocked sessions can fail during large list fetches.
        scraper = _get_idx_scraper(reset=True)
        response = scraper.get(url, params=params or {}, headers=headers or AJAIB_HEADERS, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        return response.json()


def fetch_ajaib_stock_page(
    page: int = 1,
    page_size: int = 20,
    sort_type: str = "MARKET_CAP",
    sort_direction: str = "DESC",
) -> dict:
    payload = _get_json_cloudscraper(
        AJAIB_STOCK_LIST_URL,
        params={
            "page": page,
            "page_size": page_size,
            "sort_type": sort_type,
            "sort_direction": sort_direction,
        },
        headers=AJAIB_HEADERS,
    )
    result = payload.get("result") or {}
    items = result.get("results") or []

    return {
        "count": result.get("count") or 0,
        "next": result.get("next"),
        "results": items,
        "raw": payload,
    }


def _normalize_ajaib_item(item: dict) -> dict:
    return {
        "symbol": item.get("code"),
        "name": item.get("name"),
        "price": item.get("price"),
        "icon_url": item.get("icon_url"),
        "market_cap": item.get("market_cap"),
        "volume": item.get("volume"),
        "price_1_week": item.get("price_1_week") or {},
        "price_1_month": item.get("price_1_month") or {},
    }


def find_ajaib_stock(symbol: str, page_size: int = 20) -> Tuple[Optional[dict], dict]:
    symbol = str(symbol or "").strip().upper()
    if not symbol:
        return None, {}

    first_page = fetch_ajaib_stock_page(page=1, page_size=page_size)
    count = int(first_page.get("count") or 0)
    total_pages = max(1, ceil(count / page_size))

    for item in first_page.get("results") or []:
        if str(item.get("code") or "").upper() == symbol:
            return _normalize_ajaib_item(item), first_page

    for page in range(2, total_pages + 1):
        page_payload = fetch_ajaib_stock_page(page=page, page_size=page_size)
        for item in page_payload.get("results") or []:
            if str(item.get("code") or "").upper() == symbol:
                return _normalize_ajaib_item(item), page_payload

    return None, first_page


def fetch_idx_company_profile(symbol: str) -> dict:
    params = {
        "KodeEmiten": symbol.upper(),
        "language": "id-id",
    }
    try:
        payload = _get_json(
            IDX_PROFILE_URL,
            params=params,
            headers=IDX_HEADERS,
        )
    except requests.HTTPError as exc:
        status_code = getattr(getattr(exc, "response", None), "status_code", None)
        if status_code != 403:
            raise
        payload = _get_json_cloudscraper(
            IDX_PROFILE_URL,
            params=params,
            headers=IDX_HEADERS,
        )
    except Exception:
        # Network/intermittent errors should still attempt the IDX scraper path.
        payload = _get_json_cloudscraper(
            IDX_PROFILE_URL,
            params=params,
            headers=IDX_HEADERS,
        )

    def _find_key_recursive(obj: Any, key: str) -> Any:
        if isinstance(obj, dict):
            if key in obj:
                return obj[key]
            for value in obj.values():
                found = _find_key_recursive(value, key)
                if found is not None:
                    return found
        elif isinstance(obj, list):
            for value in obj:
                found = _find_key_recursive(value, key)
                if found is not None:
                    return found
        return None

    merged: Dict[str, Any] = dict(payload) if isinstance(payload, dict) else {"raw": payload}
    profiles = _find_key_recursive(payload, "Profiles")
    profile = profiles[0] if isinstance(profiles, list) and profiles else {}

    merged["profile"] = profile

    list_keys = [
        "Direktur",
        "Komisaris",
        "Sekretaris",
        "KomiteAudit",
        "PemegangSaham",
        "AnakPerusahaan",
        "Dividen",
        "BondsAndSukuk",
        "IssuedBond",
    ]
    for key in list_keys:
        value = merged.get(key)
        if not isinstance(value, list):
            found = _find_key_recursive(payload, key)
            merged[key] = found if isinstance(found, list) else []

    return merged


def _status_to_company_status(value: Any) -> str:
    if value in (0, "0", None, ""):
        return "ACTIVE"
    return "ACTIVE"


def _build_logo_url(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    if value.startswith("http"):
        return value
    return f"https://www.idx.co.id{value}"


def _build_country_payload() -> dict:
    return {
        "name": "Indonesia",
        "code": "ID",
        "currency": "IDR",
    }


def _build_exchange_payload() -> dict:
    return {
        "code": "IDX",
        "name": "Indonesia Stock Exchange",
        "timezone": "Asia/Jakarta",
        "exchangeType": "STOCK",
    }


def _pick_ceo_from_directors(directors: Any) -> Optional[str]:
    if not isinstance(directors, list):
        return None

    president_titles = (
        "PRESIDEN DIREKTUR",
        "CEO",
        "CHIEF EXECUTIVE OFFICER",
    )

    for director in directors:
        if not isinstance(director, dict):
            continue
        name = str(director.get("Nama") or "").strip()
        title = str(director.get("Jabatan") or "").strip().upper()
        if name and any(keyword in title for keyword in president_titles):
            return name

    for director in directors:
        if not isinstance(director, dict):
            continue
        name = str(director.get("Nama") or "").strip()
        if name:
            return name

    return None


def _build_company_payload(symbol: str, ajaib_item: dict, idx_profile: dict, idx_payload: dict) -> dict:
    tanggal_pencatatan = idx_profile.get("TanggalPencatatan")
    founded_year = None
    if isinstance(tanggal_pencatatan, str) and len(tanggal_pencatatan) >= 4:
        try:
            founded_year = int(tanggal_pencatatan[:4])
        except ValueError:
            founded_year = None

    return {
        "legalName": idx_profile.get("NamaEmiten") or ajaib_item.get("name"),
        "displayName": ajaib_item.get("name") or idx_profile.get("NamaEmiten") or symbol,
        "description": idx_profile.get("KegiatanUsahaUtama"),
        "foundedYear": founded_year,
        "website": idx_profile.get("Website"),
        "logoUrl": _build_logo_url(idx_profile.get("Logo")) or ajaib_item.get("icon_url"),
        "employeeCount": None,
        "ceo": _pick_ceo_from_directors(idx_payload.get("Direktur") or []),
        "headquarters": idx_profile.get("Alamat"),
        "status": _status_to_company_status(idx_profile.get("Status")),
    }


def _build_listing_payload(symbol: str, ajaib_item: dict) -> dict:
    return {
        "symbol": symbol,
        "isin": None,
        "cusip": None,
        "assetType": "STOCK",
        "marketCap": ajaib_item.get("market_cap"),
        "price": ajaib_item.get("price"),
        "volume": ajaib_item.get("volume"),
    }


def _build_management_payload(idx_payload: dict) -> dict:
    return {
        "secretary": idx_payload.get("Sekretaris") or [],
        "directors": idx_payload.get("Direktur") or [],
        "commissioners": idx_payload.get("Komisaris") or [],
        "auditCommittee": idx_payload.get("KomiteAudit") or [],
    }


def _build_prisma_payload(symbol: str, ajaib_item: dict, idx_payload: dict) -> dict:
    profile = idx_payload.get("profile") or {}
    company = _build_company_payload(symbol, ajaib_item, profile, idx_payload)

    return {
        "symbol": symbol,
        "country": _build_country_payload(),
        "exchange": _build_exchange_payload(),
        "sector": {"name": profile.get("Sektor")},
        "subSector": {"name": profile.get("SubSektor")},
        "industry": {"name": profile.get("Industri")},
        "company": company,
        "listing": _build_listing_payload(symbol, ajaib_item),
        "market": {
            "name": ajaib_item.get("name"),
            "price": ajaib_item.get("price"),
            "market_cap": ajaib_item.get("market_cap"),
            "volume": ajaib_item.get("volume"),
            "price_1_week": ajaib_item.get("price_1_week") or {},
            "price_1_month": ajaib_item.get("price_1_month") or {},
            "icon_url": ajaib_item.get("icon_url"),
        },
        "shareholders": idx_payload.get("PemegangSaham") or [],
        "management": _build_management_payload(idx_payload),
        "subsidiaries": idx_payload.get("AnakPerusahaan") or [],
        "dividends": idx_payload.get("Dividen") or [],
        "bonds": idx_payload.get("BondsAndSukuk") or [],
        "issuedBond": idx_payload.get("IssuedBond") or [],
    }


def scrape_emiten_list(
    page: int = 1,
    page_size: int = 20,
    sort_type: str = "MARKET_CAP",
    sort_direction: str = "DESC",
) -> dict:
    ajaib_page = fetch_ajaib_stock_page(page, page_size, sort_type, sort_direction)
    enriched_items = []
    for raw_item in ajaib_page.get("results") or []:
        symbol = str(raw_item.get("code") or "").strip().upper()
        if not symbol:
            continue

        ajaib_item = _normalize_ajaib_item(raw_item)
        try:
            idx_payload = fetch_idx_company_profile(symbol)
        except Exception:
            try:
                _get_idx_scraper(reset=True)
                idx_payload = fetch_idx_company_profile(symbol)
            except Exception:
                idx_payload = {"profile": {}}

        enriched_items.append(_build_prisma_payload(symbol, ajaib_item, idx_payload))

    return {
        "mode": "list",
        "pagination": {
            "page": page,
            "page_size": page_size,
            "count": ajaib_page.get("count") or 0,
            "next": ajaib_page.get("next"),
        },
        "items": enriched_items,
    }


def scrape_emiten_detail(symbol: str) -> dict:
    ajaib_item, source_page = find_ajaib_stock(symbol)
    if not ajaib_item:
        raise ValueError(f"Stock '{symbol}' not found on Ajaib")

    idx_payload = fetch_idx_company_profile(symbol)
    prisma_payload = _build_prisma_payload(symbol, ajaib_item, idx_payload)

    return {
        "mode": "detail",
        "symbol": symbol,
        "source_page": source_page,
        "item": prisma_payload,
    }
