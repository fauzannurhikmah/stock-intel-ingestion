from app.scrapers.financial_statement import scrape_financial_statement


def fetch_and_build_financial_statement(symbol: str, year: int) -> dict:
    return scrape_financial_statement(symbol, year)
