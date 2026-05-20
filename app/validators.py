from typing import Tuple, List, Any, Optional

from utils.helper import VALID_QUARTERS


def validate_fundamental_request(body: dict) -> Tuple[str, Any, Any, List[str]]:
    body = body or {}
    symbol = (body.get("symbol") or "").strip()
    year = body.get("year")
    quarter = str(body.get("quarter", "")).strip().upper()
    quarter = quarter or None

    errors = []
    if not symbol:
        errors.append("'symbol' is required (e.g. 'BBCA')")
    if not year:
        errors.append("'year' is required (e.g. 2024)")
    else:
        try:
            year = int(year)
        except (ValueError, TypeError):
            errors.append("'year' must be a valid integer")
    if quarter and quarter not in VALID_QUARTERS:
        errors.append(f"'quarter' must be one of {sorted(VALID_QUARTERS)}")

    return symbol, year, quarter, errors


def validate_technical_request(body: dict) -> Tuple[str, List[str]]:
    body = body or {}
    emiten = (
        body.get("emiten")
        or body.get("nama_saham")
        or body.get("symbol")
        or ""
    )
    emiten = str(emiten).strip().upper()
    errors = []
    if not emiten:
        errors.append("'emiten' is required (e.g. 'BBCA')")
    return emiten, errors


def validate_emiten_request(body: dict) -> Tuple[Optional[str], int, int, str, str, List[str]]:
    body = body or {}
    symbol = str(body.get("symbol") or "").strip().upper() or None
    sort_type = str(body.get("sort_type") or "MARKET_CAP").strip().upper()
    sort_direction = str(body.get("sort_direction") or "DESC").strip().upper()

    page = body.get("page", 1)
    page_size = body.get("page_size", 20)

    errors = []
    try:
        page = int(page)
        if page < 1:
            raise ValueError
    except (TypeError, ValueError):
        errors.append("'page' must be a positive integer")
        page = 1

    try:
        page_size = int(page_size)
        if page_size < 1:
            raise ValueError
    except (TypeError, ValueError):
        errors.append("'page_size' must be a positive integer")
        page_size = 20

    if page_size > 100:
        errors.append("'page_size' must not exceed 100")

    if sort_direction not in {"ASC", "DESC"}:
        errors.append("'sort_direction' must be either 'ASC' or 'DESC'")

    return symbol, page, page_size, sort_type, sort_direction, errors


def validate_shares_data_request(body: dict) -> Tuple[Optional[str], List[str]]:
    body = body or {}
    symbol = (
        body.get("symbol")
        or body.get("emiten")
        or body.get("kode_emiten")
        or ""
    )
    symbol = str(symbol).strip().upper() or None

    errors = []
    if not symbol:
        errors.append("'symbol' is required (e.g. 'BBCA')")

    return symbol, errors
