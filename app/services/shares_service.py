from app.scraper import scrape_shares_data


def fetch_and_build_shares_data(symbol: str) -> dict:
    payload = scrape_shares_data(symbol)
    return {
        "status": "ok",
        "symbol": payload.get("symbol"),
        "count": payload.get("count") or 0,
        "shares_data": payload.get("items") or [],
    }