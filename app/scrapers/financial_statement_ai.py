from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional

from app.scrapers.common import BASE_URL, _download_file, _extract_attachment_text
from app.scrapers.fs_utilities.fs_utils import (
    _fiscal_quarter,
    _period_end_date,
    _audit_status,
    _attachment_url,
)
from app.scrapers.fs_utilities.fs_collectors import (
    fetch_financial_report_results,
    _collect_pdf_attachments,
)
from app.scrapers.fs_utilities.fs_builders import (
    _normalize_monetary_scale,
    _normalize_cash_flow_scale,
    _apply_bank_derivations,
)


def _normalize_ai_period(period: Any) -> str:
    raw_period = str(period or "").strip().upper()
    if not raw_period:
        return ""

    compact = "".join(ch for ch in raw_period if ch.isalnum())

    if any(token in compact for token in {"AUDIT", "ANNUAL", "FULL", "TAHUNAN"}):
        return "AUDIT"
    if re.search(r"(\bQ1\b|\bTW1\b|TRIWULAN\s*1|TRIWULAN\s*I|QUARTER\s*1)", raw_period):
        return "Q1"
    if re.search(r"(\bQ2\b|\bTW2\b|TRIWULAN\s*2|TRIWULAN\s*II|QUARTER\s*2)", raw_period):
        return "Q2"
    if re.search(r"(\bQ3\b|\bTW3\b|TRIWULAN\s*3|TRIWULAN\s*III|QUARTER\s*3)", raw_period):
        return "Q3"
    if re.search(r"(\bQ4\b|\bTW4\b|TRIWULAN\s*4|TRIWULAN\s*IV|QUARTER\s*4)", raw_period):
        return "Q4"

    return raw_period


def _merge_statement_item(existing: dict, incoming: dict) -> dict:
    merged = dict(existing)

    existing_confidence = float(existing.get("confidence") or 0.0)
    incoming_confidence = float(incoming.get("confidence") or 0.0)
    prefer_incoming = incoming_confidence > existing_confidence

    for key, value in incoming.items():
        if key == "confidence":
            continue
        if prefer_incoming or merged.get(key) in (None, ""):
            merged[key] = value

    merged["confidence"] = max(existing_confidence, incoming_confidence)
    return merged


def _dedupe_statement_items(items: list[dict]) -> list[dict]:
    deduped: list[dict] = []
    seen: set[tuple[str, int]] = set()

    for item in items:
        period = _normalize_ai_period(item.get("period"))
        fiscal_year = item.get("fiscalYear")
        if not period or not isinstance(fiscal_year, int):
            continue

        key = (period, fiscal_year)
        if key in seen:
            continue

        seen.add(key)
        normalized_item = dict(item)
        normalized_item["period"] = period
        normalized_item["fiscalYear"] = fiscal_year
        normalized_item["fiscalQuarter"] = _fiscal_quarter(period)
        normalized_item["periodEndDate"] = _period_end_date(fiscal_year, normalized_item["fiscalQuarter"])
        normalized_item["auditStatus"] = _audit_status(period)
        deduped.append(normalized_item)

    deduped.sort(key=lambda row: (int(row.get("fiscalYear") or 0), int(row.get("fiscalQuarter") or 99)))
    return deduped


def _derive_bank_revenue(item: dict) -> dict:
    interest_income = item.get("interestIncome")
    other_operating_income = item.get("otherNonOperatingIncome")

    if not isinstance(interest_income, (int, float)):
        return item

    if isinstance(other_operating_income, (int, float)) and other_operating_income != 0:
        item["revenue"] = float(interest_income) + float(other_operating_income)
        return item

    if item.get("revenue") in (None, 0):
        item["revenue"] = float(interest_income)

    return item


def _calculate_revenue_growth(items: list[dict]) -> list[dict]:
    lookup: dict[tuple[int, int | None], dict] = {}
    for item in items:
        fiscal_year = item.get("fiscalYear")
        fiscal_quarter = item.get("fiscalQuarter")
        if isinstance(fiscal_year, int):
            lookup[(fiscal_year, fiscal_quarter if isinstance(fiscal_quarter, int) else None)] = item

    for item in items:
        fiscal_year = item.get("fiscalYear")
        fiscal_quarter = item.get("fiscalQuarter")
        current_revenue = item.get("revenue")

        if not isinstance(fiscal_year, int) or not isinstance(fiscal_quarter, int):
            item["revenueGrowthYoY"] = None
            continue

        previous = lookup.get((fiscal_year - 1, fiscal_quarter))
        previous_revenue = previous.get("revenue") if previous else None
        if isinstance(current_revenue, (int, float)) and isinstance(previous_revenue, (int, float)) and previous_revenue not in (0, 0.0):
            item["revenueGrowthYoY"] = round((float(current_revenue) - float(previous_revenue)) / abs(float(previous_revenue)), 6)
        else:
            item["revenueGrowthYoY"] = None

    return items


def _ai_call(prompt: str) -> dict:
    from config.settings import OPENAI_API_KEY, OPENAI_MODEL
    from utils.ai import _build_client, _safe_json_parse

    if not OPENAI_API_KEY:
        return {}

    try:
        client = _build_client()
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": "You are a precise financial data extraction assistant. Return ONLY a valid JSON object matching the requested schema. Do not output markdown code fences, comments or explanations.",
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        content = response.choices[0].message.content or ""
        return _safe_json_parse(content)
    except Exception as e:
        print(f"AI call failed: {e}")
        return {}


def ai_extract_financial_statements(text: str, symbol: str, year: int, sector: str | None) -> dict:
    prompt = f"""
    You are an expert financial analyst. Extract all financial statements (income statement, balance sheet, and cash flow statement) for stock symbol '{symbol}' and fiscal year {year} from the following text document.
    
    The sector of the company is: '{sector or "general"}'.
    If the sector is 'keuangan' (financial/banking), pay close attention to banking-specific line items (e.g., interest income, interest expense, net interest income).

    Identify ALL reporting periods mentioned in the text (e.g. Q1, Q2, Q3, AUDIT/Q4). For each period and statement, extract the values.
    
    CRITICAL RULES:
    1. DETECT THE SCALE/MULTIPLIER: Detect if the numbers in the document are in millions (jutaan), thousands (ribuan), billions (miliaran), or units (satuan). You MUST scale all numeric values to full units (e.g. if currency is IDR and the table is in millions, multiply each number by 1,000,000. So "14,146,990" becomes 14146990000000). Return the final scaled numbers.
    2. PERIOD IDENTIFICATION: Identify the period ("Q1", "Q2", "Q3", "AUDIT"), fiscalYear (e.g., 2025), fiscalQuarter (1, 2, 3, or null for AUDIT), periodEndDate ("YYYY-MM-DD"), auditStatus ("AUDITED" or "UNAUDITED") for each list item.
    3. Return a float value between 0.0 and 1.0 representing your extraction 'confidence' for each list item.
    4. For banking/financial sector companies, use `otherNonOperatingIncome` for "pendapatan operasional lainnya" when present, and calculate `revenue` as `interestIncome + otherNonOperatingIncome`.
    5. If a field is not mentioned or not applicable, set its value to null.
    6. Return ONLY a valid JSON object matching the JSON schema below. Do not include markdown code fences or comments.

    JSON SCHEMA:
    {{
      "income_statement": [
        {{
          "period": "Q1" | "Q2" | "Q3" | "AUDIT",
          "fiscalYear": integer,
          "fiscalQuarter": 1 | 2 | 3 | null,
          "periodEndDate": "YYYY-MM-DD",
          "auditStatus": "AUDITED" | "UNAUDITED",
          "currency": "IDR" | "USD",
          "confidence": float,
          "revenueGrowthYoY": number | null,
          "revenue": number | null,
          "cogs": number | null,
          "grossProfit": number | null,
          "operatingExpenses": number | null,
          "sellingExpenses": number | null,
          "generalAdminExpenses": number | null,
          "rdExpenses": number | null,
          "depreciationAmort": number | null,
          "ebit": number | null,
          "ebitda": number | null,
          "operatingIncome": number | null,
          "interestExpense": number | null,
          "interestIncome": number | null,
          "otherNonOperatingIncome": number | null,
          "pretaxIncome": number | null,
          "incomeTaxExpense": number | null,
          "effectiveTaxRate": number | null,
          "netIncome": number | null,
          "netIncomeAttributable": number | null,
          "minorityInterest": number | null,
          "eps": number | null,
          "epsDiluted": number | null,
          "sharesWeightedAvg": number | null
        }}
      ],
      "balance_sheet": [
        {{
          "period": "Q1" | "Q2" | "Q3" | "AUDIT",
          "fiscalYear": integer,
          "fiscalQuarter": 1 | 2 | 3 | null,
          "periodEndDate": "YYYY-MM-DD",
          "auditStatus": "AUDITED" | "UNAUDITED",
          "currency": "IDR" | "USD",
          "confidence": float,
          "cash": number | null,
          "shortTermInvestments": number | null,
          "accountsReceivable": number | null,
          "inventory": number | null,
          "otherCurrentAssets": number | null,
          "totalCurrentAssets": number | null,
          "propertyPlantEquipment": number | null,
          "intangibleAssets": number | null,
          "goodwill": number | null,
          "longTermInvestments": number | null,
          "otherNonCurrentAssets": number | null,
          "totalNonCurrentAssets": number | null,
          "totalAssets": number | null,
          "shortTermDebt": number | null,
          "accountsPayable": number | null,
          "deferredRevenue": number | null,
          "otherCurrentLiabilities": number | null,
          "totalCurrentLiabilities": number | null,
          "longTermDebt": number | null,
          "deferredTaxLiabilities": number | null,
          "otherNonCurrentLiabilities": number | null,
          "totalNonCurrentLiabilities": number | null,
          "totalLiabilities": number | null,
          "commonStock": number | null,
          "additionalPaidInCapital": number | null,
          "retainedEarnings": number | null,
          "treasuryStock": number | null,
          "otherEquity": number | null,
          "minorityInterestEquity": number | null,
          "totalEquity": number | null,
          "bookValuePerShare": number | null,
          "netDebt": number | null,
          "workingCapital": number | null
        }}
      ],
      "cash_flow_statement": [
        {{
          "period": "Q1" | "Q2" | "Q3" | "AUDIT",
          "fiscalYear": integer,
          "fiscalQuarter": 1 | 2 | 3 | null,
          "periodEndDate": "YYYY-MM-DD",
          "auditStatus": "AUDITED" | "UNAUDITED",
          "currency": "IDR" | "USD",
          "confidence": float,
          "netIncomeStart": number | null,
          "depreciationAmort": number | null,
          "stockBasedCompensation": number | null,
          "changeInWorkingCapital": number | null,
          "changeInReceivables": number | null,
          "changeInInventory": number | null,
          "changeInPayables": number | null,
          "otherOperatingActivities": number | null,
          "netCashFromOperations": number | null,
          "capitalExpenditures": number | null,
          "acquisitions": number | null,
          "purchaseOfInvestments": number | null,
          "saleOfInvestments": number | null,
          "otherInvestingActivities": number | null,
          "netCashFromInvesting": number | null,
          "debtIssuance": number | null,
          "debtRepayment": number | null,
          "commonStockIssuance": number | null,
          "commonStockRepurchase": number | null,
          "dividendsPaid": number | null,
          "otherFinancingActivities": number | null,
          "netCashFromFinancing": number | null,
          "netChangeInCash": number | null,
          "cashBeginningPeriod": number | null,
          "cashEndPeriod": number | null,
          "freeCashFlow": number | null
        }}
      ]
    }}

    TEXT DOCUMENT:
    {text[:45000]}
    """
    return _ai_call(prompt)


def _adjust_cumulative_quarter_items(items: list[dict]) -> list[dict]:
    numeric_fields = [
        "revenue", "cogs", "grossProfit", "operatingExpenses", "sellingExpenses",
        "generalAdminExpenses", "rdExpenses", "depreciationAmort", "ebit", "ebitda",
        "operatingIncome", "interestExpense", "interestIncome", "otherNonOperatingIncome",
        "pretaxIncome", "incomeTaxExpense", "netIncome", "netIncomeAttributable",
        "minorityInterest", "eps", "epsDiluted", "sharesWeightedAvg",
    ]
    original_items = [dict(item) for item in items]

    for index, item in enumerate(items):
        quarter = item.get("fiscalQuarter")
        if quarter not in (2, 3):
            continue

        previous = None
        for prior in reversed(original_items[:index]):
            if (
                prior.get("fiscalYear") == item.get("fiscalYear")
                and prior.get("fiscalQuarter") == quarter - 1
            ):
                previous = prior
                break

        if not previous:
            continue

        for field in numeric_fields:
            current = item.get(field)
            previous_value = previous.get(field)
            if isinstance(current, (int, float)) and isinstance(previous_value, (int, float)):
                item[field] = current - previous_value

        pretax = item.get("pretaxIncome")
        tax = item.get("incomeTaxExpense")
        if isinstance(pretax, (int, float)) and isinstance(tax, (int, float)) and pretax != 0:
            item["effectiveTaxRate"] = round(abs(float(tax)) / abs(float(pretax)), 6)
        else:
            item["effectiveTaxRate"] = None

    return items


def scrape_financial_statement_ai(symbol: str, year: int, sector: Optional[str] = None) -> dict:
    symbol = symbol.upper()
    results = fetch_financial_report_results(symbol, year)
    report_attachments = _collect_pdf_attachments(results)

    income_items: list[dict] = []
    balance_items: list[dict] = []
    cash_flow_items: list[dict] = []
    seen_income_period_year: set[tuple[str, int]] = set()
    seen_balance_period_year: set[tuple[str, int]] = set()
    seen_cash_flow_period_year: set[tuple[str, int]] = set()

    for result, attachment in report_attachments:
        file_name = str(attachment.get("File_Name") or attachment.get("file_name") or "")
        file_url = _attachment_url(attachment)
        if not file_url:
            continue

        try:
            content = _download_file(file_url)
            extracted_text = _extract_attachment_text(file_name, content)

            ai_data = ai_extract_financial_statements(extracted_text, symbol, year, sector)

            income_list = ai_data.get("income_statement") or []
            balance_list = ai_data.get("balance_sheet") or []
            cash_flow_list = ai_data.get("cash_flow_statement") or []

            for item in income_list:
                period = _normalize_ai_period(item.get("period"))
                f_year = item.get("fiscalYear") or year
                if not period or not f_year:
                    continue
                dedup_key = (str(period), int(f_year))
                if dedup_key not in seen_income_period_year:
                    seen_income_period_year.add(dedup_key)
                    # Normalize metadata fields
                    item["period"] = period
                    item["fiscalYear"] = int(f_year)
                    item["fiscalQuarter"] = _fiscal_quarter(period)
                    item["periodEndDate"] = _period_end_date(int(f_year), item["fiscalQuarter"])
                    item["auditStatus"] = _audit_status(period)
                    item["confidence"] = float(item.get("confidence") or 0.0)
                    # Apply safety derivations
                    item = _normalize_monetary_scale(item)
                    item = _apply_bank_derivations(item)
                    item = _derive_bank_revenue(item)
                    income_items.append(item)
                else:
                    for index, existing in enumerate(income_items):
                        existing_key = (
                            str(existing.get("period") or ""),
                            int(existing.get("fiscalYear") or 0),
                        )
                        if existing_key == dedup_key:
                            item["period"] = period
                            item["fiscalYear"] = int(f_year)
                            item["fiscalQuarter"] = _fiscal_quarter(period)
                            item["periodEndDate"] = _period_end_date(int(f_year), item["fiscalQuarter"])
                            item["auditStatus"] = _audit_status(period)
                            item["confidence"] = float(item.get("confidence") or 0.0)
                            item = _normalize_monetary_scale(item)
                            item = _apply_bank_derivations(item)
                            item = _derive_bank_revenue(item)
                            income_items[index] = _merge_statement_item(existing, item)
                            break

            for item in balance_list:
                period = _normalize_ai_period(item.get("period"))
                f_year = item.get("fiscalYear") or year
                if not period or not f_year:
                    continue
                dedup_key = (str(period), int(f_year))
                if dedup_key not in seen_balance_period_year:
                    seen_balance_period_year.add(dedup_key)
                    # Normalize metadata fields
                    item["period"] = period
                    item["fiscalYear"] = int(f_year)
                    item["fiscalQuarter"] = _fiscal_quarter(period)
                    item["periodEndDate"] = _period_end_date(int(f_year), item["fiscalQuarter"])
                    item["auditStatus"] = _audit_status(period)
                    item["confidence"] = float(item.get("confidence") or 0.0)
                    item = _normalize_monetary_scale(item)
                    balance_items.append(item)
                else:
                    for index, existing in enumerate(balance_items):
                        existing_key = (
                            str(existing.get("period") or ""),
                            int(existing.get("fiscalYear") or 0),
                        )
                        if existing_key == dedup_key:
                            item["period"] = period
                            item["fiscalYear"] = int(f_year)
                            item["fiscalQuarter"] = _fiscal_quarter(period)
                            item["periodEndDate"] = _period_end_date(int(f_year), item["fiscalQuarter"])
                            item["auditStatus"] = _audit_status(period)
                            item["confidence"] = float(item.get("confidence") or 0.0)
                            item = _normalize_monetary_scale(item)
                            balance_items[index] = _merge_statement_item(existing, item)
                            break

            for item in cash_flow_list:
                period = _normalize_ai_period(item.get("period"))
                f_year = item.get("fiscalYear") or year
                if not period or not f_year:
                    continue
                dedup_key = (str(period), int(f_year))
                if dedup_key not in seen_cash_flow_period_year:
                    seen_cash_flow_period_year.add(dedup_key)
                    # Normalize metadata fields
                    item["period"] = period
                    item["fiscalYear"] = int(f_year)
                    item["fiscalQuarter"] = _fiscal_quarter(period)
                    item["periodEndDate"] = _period_end_date(int(f_year), item["fiscalQuarter"])
                    item["auditStatus"] = _audit_status(period)
                    item["confidence"] = float(item.get("confidence") or 0.0)
                    item = _normalize_cash_flow_scale(item)
                    cash_flow_items.append(item)
                else:
                    for index, existing in enumerate(cash_flow_items):
                        existing_key = (
                            str(existing.get("period") or ""),
                            int(existing.get("fiscalYear") or 0),
                        )
                        if existing_key == dedup_key:
                            item["period"] = period
                            item["fiscalYear"] = int(f_year)
                            item["fiscalQuarter"] = _fiscal_quarter(period)
                            item["periodEndDate"] = _period_end_date(int(f_year), item["fiscalQuarter"])
                            item["auditStatus"] = _audit_status(period)
                            item["confidence"] = float(item.get("confidence") or 0.0)
                            item = _normalize_cash_flow_scale(item)
                            cash_flow_items[index] = _merge_statement_item(existing, item)
                            break

        except Exception as e:
            print(f"Failed to process attachment {file_name}: {e}")
            continue

    # Sort
    income_items = _dedupe_statement_items(income_items)
    balance_items = _dedupe_statement_items(balance_items)
    cash_flow_items = _dedupe_statement_items(cash_flow_items)

    # Adjust cumulative quarter values for the income statement
    income_items = _adjust_cumulative_quarter_items(income_items)
    income_items = _calculate_revenue_growth(income_items)

    income_items = [item for item in income_items if int(item.get("fiscalYear") or 0) == year]

    return {
        "status": "ok",
        "symbol": symbol.upper(),
        "year": year,
        "income_statement": {
            "count": len(income_items),
            "items": income_items,
        },
        "balance_sheet": {
            "count": len(balance_items),
            "items": balance_items,
        },
        "cash_flow_statement": {
            "count": len(cash_flow_items),
            "items": cash_flow_items,
        },
    }
