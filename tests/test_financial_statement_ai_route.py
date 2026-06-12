import unittest
from unittest.mock import patch
from app import create_app
from app.scrapers import financial_statement_ai as ai_module


def _build_mock_ai_response():
    return {
        "income_statement": [
            {
                "period": "Q1",
                "fiscalYear": 2024,
                "fiscalQuarter": 1,
                "periodEndDate": "2024-03-31",
                "auditStatus": "UNAUDITED",
                "currency": "IDR",
                "confidence": 0.84,
                "revenue": 999999999,
                "cogs": 3248158000000,
                "grossProfit": None,
                "operatingExpenses": 9637633000000,
                "sellingExpenses": None,
                "generalAdminExpenses": None,
                "rdExpenses": None,
                "depreciationAmort": None,
                "ebit": None,
                "ebitda": None,
                "operatingIncome": 21118560000000,
                "interestExpense": 3248158000000,
                "interestIncome": 22963761000000,
                "otherNonOperatingIncome": 6451724000000,
                "pretaxIncome": 15915029000000,
                "incomeTaxExpense": 3036522000000,
                "effectiveTaxRate": 0.19,
                "netIncome": 12878507000000,
                "netIncomeAttributable": 12879486000000,
                "minorityInterest": -979000000,
                "eps": 104,
                "epsDiluted": None,
                "sharesWeightedAvg": None
            },
            {
                "period": "Q1",
                "fiscalYear": 2025,
                "fiscalQuarter": 1,
                "periodEndDate": "2025-03-31",
                "auditStatus": "UNAUDITED",
                "currency": "IDR",
                "confidence": 0.95,
                "revenue": 24366718000000,
                "cogs": 3248158000000,
                "grossProfit": None,
                "operatingExpenses": 9637633000000,
                "sellingExpenses": None,
                "generalAdminExpenses": None,
                "rdExpenses": None,
                "depreciationAmort": None,
                "ebit": None,
                "ebitda": None,
                "operatingIncome": 21118560000000,
                "interestExpense": 3248158000000,
                "interestIncome": 24366718000000,
                "otherNonOperatingIncome": 7005767000000,
                "pretaxIncome": 17455662000000,
                "incomeTaxExpense": 3308672000000,
                "effectiveTaxRate": 0.19,
                "netIncome": 14146990000000,
                "netIncomeAttributable": 14146131000000,
                "minorityInterest": 859000000,
                "eps": 115,
                "epsDiluted": None,
                "sharesWeightedAvg": None
            }
        ],
        "balance_sheet": [
            {
                "period": "Q1",
                "fiscalYear": 2025,
                "fiscalQuarter": 1,
                "periodEndDate": "2025-03-31",
                "auditStatus": "UNAUDITED",
                "currency": "IDR",
                "confidence": 0.92,
                "cash": 28032494000000,
                "shortTermInvestments": 56182969000000,
                "accountsReceivable": None,
                "inventory": None,
                "otherCurrentAssets": None,
                "totalCurrentAssets": None,
                "propertyPlantEquipment": None,
                "intangibleAssets": None,
                "goodwill": None,
                "longTermInvestments": None,
                "otherNonCurrentAssets": None,
                "totalNonCurrentAssets": None,
                "totalAssets": 1533763445000000,
                "shortTermDebt": None,
                "accountsPayable": None,
                "deferredRevenue": None,
                "otherCurrentLiabilities": None,
                "totalCurrentLiabilities": None,
                "longTermDebt": None,
                "deferredTaxLiabilities": None,
                "otherNonCurrentLiabilities": None,
                "totalNonCurrentLiabilities": None,
                "totalLiabilities": 1278027110000000,
                "commonStock": None,
                "additionalPaidInCapital": None,
                "retainedEarnings": None,
                "treasuryStock": None,
                "otherEquity": None,
                "minorityInterestEquity": None,
                "totalEquity": 246520509000000,
                "bookValuePerShare": None,
                "netDebt": None,
                "workingCapital": None
            }
        ],
        "cash_flow_statement": [
            {
                "period": "Q1",
                "fiscalYear": 2025,
                "fiscalQuarter": 1,
                "periodEndDate": "2025-03-31",
                "auditStatus": "UNAUDITED",
                "currency": "IDR",
                "confidence": 0.94,
                "netIncomeStart": 14146990000000,
                "depreciationAmort": None,
                "stockBasedCompensation": None,
                "changeInWorkingCapital": None,
                "changeInReceivables": None,
                "changeInInventory": None,
                "changeInPayables": None,
                "otherOperatingActivities": None,
                "netCashFromOperations": 35183351000000,
                "capitalExpenditures": 25981888000000,
                "acquisitions": None,
                "purchaseOfInvestments": None,
                "saleOfInvestments": None,
                "otherInvestingActivities": None,
                "netCashFromInvesting": -25981888000000,
                "debtIssuance": None,
                "debtRepayment": None,
                "commonStockIssuance": None,
                "commonStockRepurchase": None,
                "dividendsPaid": 9307341000000,
                "otherFinancingActivities": None,
                "netCashFromFinancing": -9307341000000,
                "netChangeInCash": -105878000000,
                "cashBeginningPeriod": 33254736000000,
                "cashEndPeriod": 33148858000000,
                "freeCashFlow": None
            }
        ]
    }


class TestFinancialStatementAiRoute(unittest.TestCase):
    def setUp(self):
        self.app = create_app()
        self.client = self.app.test_client()

    def test_financial_statement_ai_route_requires_symbol_and_year(self):
        response = self.client.get("/api/financial-statement-ai")
        self.assertEqual(response.status_code, 400)
        body = response.get_json()
        self.assertEqual(body["status"], "error")
        self.assertTrue(any("symbol" in error for error in body["errors"]))
        self.assertTrue(any("year" in error for error in body["errors"]))

    @patch("app.scrapers.financial_statement_ai.fetch_financial_report_results")
    @patch("app.scrapers.financial_statement_ai._download_file")
    @patch("app.scrapers.financial_statement_ai._extract_attachment_text")
    @patch("app.scrapers.financial_statement_ai.ai_extract_financial_statements")
    def test_scrape_financial_statement_ai_extracts_values_with_confidence(
        self,
        mock_ai_extract,
        mock_extract,
        mock_download,
        mock_fetch
    ):
        mock_fetch.return_value = [
            {
                "Report_Period": "TW1",
                "Report_Year": "2025",
                "Attachments": [
                    {
                        "File_Name": "FinancialStatement-2025-I-BBRI.pdf",
                        "File_Path": "/fake/report.pdf",
                        "File_Type": ".pdf",
                    }
                ],
            }
        ]
        mock_download.return_value = b"%PDF-1.4 fake bytes"
        mock_extract.return_value = "fake extracted text"
        mock_ai_extract.return_value = _build_mock_ai_response()

        payload = ai_module.scrape_financial_statement_ai("bbri", 2025)

        self.assertEqual(payload["status"], "ok")
        self.assertEqual(payload["symbol"], "BBRI")
        self.assertEqual(payload["year"], 2025)

        # Income Statement tests
        self.assertEqual(payload["income_statement"]["count"], 1)
        income_item = payload["income_statement"]["items"][0]
        self.assertEqual(income_item["revenue"], 31372485000000)
        self.assertEqual(income_item["netIncome"], 14146990000000)
        self.assertEqual(income_item["confidence"], 0.95)
        self.assertEqual(income_item["period"], "Q1")
        self.assertAlmostEqual(income_item["revenueGrowthYoY"], 0.06653, places=5)

        # Balance Sheet tests
        self.assertEqual(payload["balance_sheet"]["count"], 1)
        balance_item = payload["balance_sheet"]["items"][0]
        self.assertEqual(balance_item["cash"], 28032494000000)
        self.assertEqual(balance_item["totalAssets"], 1533763445000000)
        self.assertEqual(balance_item["confidence"], 0.92)

        # Cash Flow tests
        self.assertEqual(payload["cash_flow_statement"]["count"], 1)
        cash_flow_item = payload["cash_flow_statement"]["items"][0]
        self.assertEqual(cash_flow_item["netCashFromOperations"], 35183351000000)
        self.assertEqual(cash_flow_item["netCashFromInvesting"], -25981888000000)
        self.assertEqual(cash_flow_item["confidence"], 0.94)

    @patch("app.scrapers.financial_statement_ai.fetch_financial_report_results")
    @patch("app.scrapers.financial_statement_ai._download_file")
    @patch("app.scrapers.financial_statement_ai._extract_attachment_text")
    @patch("app.scrapers.financial_statement_ai.ai_extract_financial_statements")
    def test_scrape_financial_statement_ai_dedupes_normalized_periods(
        self,
        mock_ai_extract,
        mock_extract,
        mock_download,
        mock_fetch
    ):
        mock_fetch.return_value = [
            {
                "Report_Period": "TW2",
                "Report_Year": "2025",
                "Attachments": [
                    {
                        "File_Name": "FinancialStatement-2025-II-BBCA.pdf",
                        "File_Path": "/fake/report.pdf",
                        "File_Type": ".pdf",
                    }
                ],
            }
        ]
        mock_download.return_value = b"%PDF-1.4 fake bytes"
        mock_extract.return_value = "fake extracted text"
        mock_ai_extract.return_value = {
            "income_statement": [
                {
                    "period": "TW2",
                    "fiscalYear": 2025,
                    "fiscalQuarter": 2,
                    "periodEndDate": "2025-06-30",
                    "auditStatus": "UNAUDITED",
                    "currency": "IDR",
                    "confidence": 0.81,
                    "revenue": 100,
                    "netIncome": 10,
                },
                {
                    "period": "Q2",
                    "fiscalYear": 2025,
                    "fiscalQuarter": 2,
                    "periodEndDate": "2025-06-30",
                    "auditStatus": "UNAUDITED",
                    "currency": "IDR",
                    "confidence": 0.93,
                    "revenue": 200,
                    "netIncome": 20,
                },
            ],
            "balance_sheet": [],
            "cash_flow_statement": [],
        }

        payload = ai_module.scrape_financial_statement_ai("bbca", 2025, sector="keuangan")

        self.assertEqual(payload["income_statement"]["count"], 1)
        income_item = payload["income_statement"]["items"][0]
        self.assertEqual(income_item["period"], "Q2")
        self.assertEqual(income_item["confidence"], 0.93)
        self.assertEqual(income_item["revenue"], 200)
        self.assertEqual(income_item["netIncome"], 20)


if __name__ == "__main__":
    unittest.main()
