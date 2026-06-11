from typing import Tuple, List, Any, Optional

def validate_financial_statement_ai_request(body: dict) -> Tuple[Optional[str], Any, Optional[str], List[str]]:
    body = body or {}
    symbol = (
        body.get("symbol")
        or body.get("emiten")
        or body.get("kode_emiten")
        or ""
    )
    symbol = str(symbol).strip().upper() or None
    year = body.get("year")
    sector = str(body.get("sector") or "").strip().lower() or None

    errors = []
    if not symbol:
        errors.append("'symbol' is required (e.g. 'BBCA')")
    if not year:
        errors.append("'year' is required (e.g. 2025)")
    else:
        try:
            year = int(year)
        except (TypeError, ValueError):
            errors.append("'year' must be a valid integer")

    return symbol, year, sector, errors
