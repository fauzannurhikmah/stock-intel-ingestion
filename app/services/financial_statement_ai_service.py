from app.scrapers.financial_statement_ai import scrape_financial_statement_ai

def fetch_and_build_financial_statement_ai(symbol: str, year: int, sector: str | None = None) -> dict:
    """Fetch and build financial statement using AI scraper"""
    return scrape_financial_statement_ai(symbol, year, sector)
