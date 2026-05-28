from app import create_app
import app.scraper as scraper_module


def test_fetch_idx_stock_price_normalizes_payload(monkeypatch):
    captured = {}

    def fake_get_json_cloudscraper(url, params=None, headers=None):
        captured["url"] = url
        captured["params"] = params
        captured["headers"] = headers
        return {
            "KodeEmiten": "bbca",
            "Replies": [{"No": 1, "StockCode": "BBCA"}],
        }

    monkeypatch.setattr(scraper_module, "_get_json_cloudscraper", fake_get_json_cloudscraper)

    payload = scraper_module.fetch_idx_stock_price("bbca")

    assert captured["url"].endswith("GetTradingInfoSS")
    assert captured["params"] == {"code": "BBCA", "start": 0, "length": 2000}
    assert payload["KodeEmiten"] == "BBCA"
    assert payload["replies"] == [{"No": 1, "StockCode": "BBCA"}]


def test_stock_price_route_requires_symbol():
    app = create_app()
    client = app.test_client()

    response = client.get("/api/stock-price")

    assert response.status_code == 400
    body = response.get_json()
    assert body["status"] == "error"
    assert any("symbol" in error for error in body["errors"])


def test_stock_price_route_returns_payload(monkeypatch):
    def fake_fetch_and_build_stock_price(symbol):
        return {
            "KodeEmiten": symbol,
            "replies": [{"No": 1, "StockCode": symbol}],
        }

    monkeypatch.setattr("app.routes.fetch_and_build_stock_price", fake_fetch_and_build_stock_price)

    app = create_app()
    client = app.test_client()

    response = client.get("/api/stock-price?symbol=bbca")

    assert response.status_code == 200
    body = response.get_json()
    assert body["KodeEmiten"] == "BBCA"
    assert body["replies"][0]["StockCode"] == "BBCA"


def test_ajaib_stock_market_route_returns_payload(monkeypatch):
    def fake_fetch_ajaib_stock_market():
        return {
            "err_code": "EC0000000",
            "err_message": "APPROVED/OK",
            "result": {
                "count": 1,
                "results": [
                    {
                        "code": "BBCA",
                        "name": "Bank Central Asia Tbk.",
                        "price": 6100,
                    }
                ],
            },
        }

    monkeypatch.setattr("app.routes.fetch_ajaib_stock_market", fake_fetch_ajaib_stock_market)

    app = create_app()
    client = app.test_client()

    response = client.get("/api/ajaib-stock-market")

    assert response.status_code == 200
    body = response.get_json()
    assert body["err_code"] == "EC0000000"
    assert body["result"]["count"] == 1
    assert body["result"]["results"][0]["code"] == "BBCA"