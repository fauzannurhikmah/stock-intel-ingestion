from __future__ import annotations

from io import BytesIO
import re
from typing import Any

from openpyxl import load_workbook

from app.scrapers.common import BASE_URL, _download_file, _get


SUPPORTED_EXTENSIONS = {".xlsx", ".xls"}
MAX_ATTACHMENTS_TO_PARSE = 20

NUMERIC_ALIASES = {
    "revenue": [
        "jumlahpendapatanoperasional",
        "totalpendapatanoperasional",
        "pendapatanoperasional",
        "revenue",
        "pendapatan",
        "penjualan",
        "sales",
        "totalrevenue",
        "netsales",
    ],
    "cogs": ["cogs", "costofgoodsold", "costofsales", "costofrevenue", "bebanpokokpenjualan", "bebanpokok"],
    "grossProfit": ["grossprofit", "labakotor"],
    "operatingExpenses": ["operatingexpenses", "operatingexpense", "bebanoperasional", "bebanoperasi"],
    "sellingExpenses": ["sellingexpenses", "bebanpenjualan"],
    "generalAdminExpenses": ["generaladminexpenses", "generaladministrativeexpenses", "bebanumumdanadministrasi", "bebanadministrasi"],
    "rdExpenses": ["rdexpenses", "researchdevelopmentexpenses", "researchanddevelopment"],
    "depreciationAmort": ["depreciationamort", "depreciationandamortization", "depreciationamortization", "penyusutandanamortisasi"],
    "ebit": ["ebit", "earningbeforeinterestandtax"],
    "ebitda": ["ebitda"],
    "operatingIncome": ["operatingincome", "operatingprofit", "labaoperasional", "labaoperasi", "labausaha"],
    "interestExpense": ["interestexpense", "bebanbunga", "financecost"],
    "interestIncome": ["interestincome", "pendapatanbunga"],
    "otherNonOperatingIncome": ["othernonoperatingincome", "otherincome", "nonoperatingincome", "pendapatanlainlain"],
    "pretaxIncome": ["pretaxincome", "profitbeforetax", "incomebeforetax", "labasebelumpajak"],
    "incomeTaxExpense": [
        "bebanpajakpenghasilan",
        "manfaatbebanpajakpenghasilan",
        "incometaxexpense",
        "taxexpense",
    ],
    "effectiveTaxRate": ["effectivetaxrate", "taxrate", "tarifpajakefektif"],
    "netIncome": ["netincome", "netprofit", "lababersih", "labaperiodeberjalan", "jumlahlaba"],
    "netIncomeAttributable": [
        "netincomeattributable",
        "attributabletoowners",
        "labaatribusikepadapemilikentitasinduk",
        "labarugiyangdapatdiatribusikankepadapemilikentitasinduk",
        "labayangdapatdiatribusikankepadapemilikentitasinduk",
    ],
    "minorityInterest": ["minorityinterest", "noncontrollinginterest", "kepentingannonpengendali"],
    "eps": ["eps", "earningpershare", "labapersahamdasar"],
    "epsDiluted": ["epsdiluted", "dilutedeps", "labapersahamdilusian"],
    "sharesWeightedAvg": ["weightedaverageshares", "sharesweightedavg", "rataratasahamberedartertimbang"],
}

BALANCE_NUMERIC_ALIASES = {
    "cash": ["cash", "kasdansetarakas", "kasdanbank", "cashandcashequivalents"],
    "shortTermInvestments": ["shortterminvestments", "investasijangkapendek", "marketablesecurities"],
    "accountsReceivable": ["accountsreceivable", "piutangusaha", "tradeandotherreceivables"],
    "inventory": ["inventory", "persediaan"],
    "otherCurrentAssets": ["othercurrentassets", "asetlancarlainnya", "asetlancarllain"],
    "totalCurrentAssets": ["totalcurrentassets", "jumlahasetlancar", "asetlancar"],
    "propertyPlantEquipment": ["propertyplantequipment", "asettetap", "fixedassets"],
    "intangibleAssets": ["intangibleassets", "asettakberwujud", "asettidakberwujud"],
    "goodwill": ["goodwill"],
    "longTermInvestments": ["longterminvestments", "investasijangkapanjang"],
    "otherNonCurrentAssets": ["othernoncurrentassets", "asettidaklancarlainnya", "asetnonlancarlainnya"],
    "totalNonCurrentAssets": ["totalnoncurrentassets", "jumlahasettidaklancar", "asettidaklancar"],
    "totalAssets": ["totalassets", "jumlahaset", "totalaset"],
    "shortTermDebt": ["shorttermdebt", "utangjangkapendek", "pinjamanjangkapendek", "liabilitasjangkapendekberbunga"],
    "accountsPayable": ["accountspayable", "utangusaha", "tradepayables"],
    "deferredRevenue": ["deferredrevenue", "pendapatanditerimadimuka"],
    "otherCurrentLiabilities": ["othercurrentliabilities", "liabilitaslancarlainnya", "utanglancarlainnya"],
    "totalCurrentLiabilities": ["totalcurrentliabilities", "jumlahlibilitaslancar", "liabilitaslancar"],
    "longTermDebt": ["longtermdebt", "utangjangkapanjang", "pinjamanjangkapanjang", "liabilitasjangkapanjangberbunga"],
    "deferredTaxLiabilities": ["deferredtaxliabilities", "liabilitaspajaktangguhan"],
    "otherNonCurrentLiabilities": ["othernoncurrentliabilities", "liabilitasjangkapanjanglainnya", "liabilitasnonlancarlainnya"],
    "totalNonCurrentLiabilities": ["totalnoncurrentliabilities", "jumlahliabilitasjangkapanjang", "liabilitasnonlancar"],
    "totalLiabilities": ["totalliabilities", "jumlahliabilitas", "totalliabilitas"],
    "commonStock": ["commonstock", "modaldisetor", "modal saham", "issuedandfullypaidcapital"],
    "additionalPaidInCapital": ["additionalpaidincapital", "tambahanmodaldisetor", "agio"],
    "retainedEarnings": ["retainedearnings", "saldo laba", "labaditahan"],
    "treasuryStock": ["treasurystock", "sahamtreasury", "sahamdiperoleh kembali"],
    "otherEquity": ["otherequity", "ekuitaslainnya"],
    "minorityInterestEquity": ["minorityinterestequity", "kepentingannonpengendali", "noncontrollinginterest"],
    "totalEquity": ["totalequity", "jumlahekuitas", "totalekuitas"],
}

FIELD_BLOCKED_TOKENS = {
    "incomeTaxExpense": ["sebelumpajak", "beforetax"],
}

TEXT_BLOCKLIST = [
    "komprehensif lain",
    "catatan",
    "note",
    "explanation",
    "penjelasan",
    "change in name",
]


def _normalize_text(value: Any) -> str:
    text = str(value or "").strip().lower()
    return "".join(ch for ch in text if ch.isalnum())


def _normalize_file_extension(file_name: str, file_type: str | None = None) -> str:
    ext = str(file_type or "").strip().lower()
    if ext and not ext.startswith("."):
        ext = f".{ext}"
    if ext in SUPPORTED_EXTENSIONS:
        return ext

    lower_name = str(file_name or "").lower()
    if "." in lower_name:
        guessed = f".{lower_name.rsplit('.', 1)[-1]}"
        if guessed in SUPPORTED_EXTENSIONS:
            return guessed
    return ""


def _to_number(value: Any):
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    # Guard against narrative/text paragraphs being parsed as numbers.
    if len(text) > 64 and not re.search(r"\d", text):
        return None

    negative = text.startswith("(") and text.endswith(")")
    text = text.replace("(", "").replace(")", "")

    text = text.replace(" ", "")
    text = text.replace("Rp", "").replace("IDR", "").replace("idr", "")
    text = text.replace("USD", "").replace("usd", "")
    text = text.replace("%", "")

    multiplier = 1
    lower_text = text.lower()
    if "triliun" in lower_text:
        multiplier = 1_000_000_000_000
    elif "miliar" in lower_text:
        multiplier = 1_000_000_000
    elif "juta" in lower_text:
        multiplier = 1_000_000

    text = re.sub(r"[^0-9,.-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None

    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif text.count(",") > 1 and "." not in text:
        text = text.replace(",", "")
    elif text.count(".") > 1 and "," not in text:
        text = text.replace(".", "")
    elif "," in text:
        tail = text.split(",")[-1]
        if len(tail) == 3 and text.replace(",", "").replace("-", "").isdigit():
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")

    try:
        number = float(text) * multiplier
    except ValueError:
        return None

    if negative:
        number = -number

    return number


def _detect_unit_multiplier(rows: list[list[Any]]) -> float:
    for row in rows[:100]:
        for cell in row[:20]:
            text = str(cell or "").strip().lower()
            if not text:
                continue
            if "dalam jutaan" in text:
                return 1_000_000.0
            if "dalam miliaran" in text or "dalam miliar" in text:
                return 1_000_000_000.0
            if "dalam ribuan" in text:
                return 1_000.0
            if "dalam triliunan" in text or "dalam triliun" in text:
                return 1_000_000_000_000.0
            if "million" in text and "rupiah" in text:
                return 1_000_000.0
            if "billion" in text and "rupiah" in text:
                return 1_000_000_000.0
    return 1.0


def _to_ratio(value: Any):
    number = _to_number(value)
    if number is None:
        return None
    if abs(number) > 1:
        return round(number / 100.0, 6)
    return round(number, 6)


def _attachment_url(attachment: dict) -> str:
    file_path = str(attachment.get("File_Path") or attachment.get("file_path") or "").strip()
    if not file_path:
        return ""
    if file_path.startswith("http"):
        return file_path
    return f"{BASE_URL}{file_path}"


def _score_attachment(attachment: dict) -> int:
    file_name = str(attachment.get("File_Name") or attachment.get("file_name") or "").lower()
    score = 0
    for keyword in ["financialstatement", "financial statement", "laporankeuangan", "laporan keuangan", "laba rugi", "income statement"]:
        if keyword in file_name:
            score += 4
    if file_name.endswith(".xlsx"):
        score += 2
    if "xbrl" in file_name:
        score += 1
    return score


def _is_spreadsheet_attachment(attachment: dict) -> bool:
    file_name = str(attachment.get("File_Name") or attachment.get("file_name") or "")
    file_type = str(attachment.get("File_Type") or attachment.get("file_type") or "")
    ext = _normalize_file_extension(file_name, file_type)
    return ext in SUPPORTED_EXTENSIONS


def fetch_financial_report_results(symbol: str, year: int) -> list[dict]:
    url = f"{BASE_URL}/primary/ListedCompany/GetFinancialReport"
    params = {
        "periode": "*",
        "year": str(year),
        "indexFrom": 0,
        "pageSize": 1000,
        "reportType": "rdf",
        "kodeEmiten": symbol.upper(),
    }
    response_data = _get(url, params)
    results = response_data.get("Results") or response_data.get("results") or []
    return results if isinstance(results, list) else []


def _collect_spreadsheet_attachments(results: list[dict]) -> list[tuple[dict, dict]]:
    collected: list[tuple[dict, dict]] = []
    for result in results:
        attachments = result.get("Attachments") or result.get("attachments") or []
        if not isinstance(attachments, list):
            continue
        for attachment in attachments:
            if isinstance(attachment, dict) and _is_spreadsheet_attachment(attachment):
                collected.append((result, attachment))

    collected.sort(key=lambda item: _score_attachment(item[1]), reverse=True)
    return collected[:MAX_ATTACHMENTS_TO_PARSE]


def _sheet_score(sheet_name: str) -> int:
    lower = (sheet_name or "").lower()
    score = 0
    for keyword in ["income", "laba", "rugi", "statement", "komprehensif", "profit"]:
        if keyword in lower:
            score += 3
    return score


def _balance_sheet_score(sheet_name: str) -> int:
    lower = (sheet_name or "").lower()
    score = 0
    for keyword in ["balance", "position", "neraca", "aset", "liabilitas", "ekuitas"]:
        if keyword in lower:
            score += 3
    return score


def _normalize_rows(sheet) -> list[list[Any]]:
    rows = []
    for row in sheet.iter_rows(values_only=True):
        values = list(row)
        if any(cell not in (None, "") for cell in values):
            rows.append(values)
    return rows


def _find_year_columns(rows: list[list[Any]], year: int) -> dict[int, int]:
    year_columns: dict[int, int] = {}
    targets = {year, year - 1, year - 2}

    for row in rows[:20]:
        for idx, cell in enumerate(row):
            text = str(cell or "").strip()
            match = re.search(r"(19|20)\d{2}", text)
            if not match:
                continue
            year_value = int(match.group(0))
            if year_value in targets and year_value not in year_columns:
                year_columns[year_value] = idx

    return year_columns


def _row_matches_alias(row: list[Any], aliases: list[str]) -> tuple[int, str] | None:
    normalized_aliases = [_normalize_text(alias) for alias in aliases]
    for index, cell in enumerate(row[:4]):
        normalized_cell = _normalize_text(cell)
        if not normalized_cell:
            continue
        raw_label = str(cell or "").strip().lower()
        if any(blocked in raw_label for blocked in TEXT_BLOCKLIST):
            continue
        if any(alias and alias in normalized_cell for alias in normalized_aliases):
            return index, str(cell).strip()
    return None


def _extract_numeric_value(row: list[Any], label_index: int, year_col: int | None, unit_multiplier: float):
    candidates: list[Any] = []

    if year_col is not None and 0 <= year_col < len(row):
        candidates.append(row[year_col])

    # Fallback to cells right to the label, usually where values are placed.
    for i in range(label_index + 1, min(len(row), label_index + 6)):
        candidates.append(row[i])

    for candidate in candidates:
        number = _to_number(candidate)
        if number is not None:
            return number * unit_multiplier

    return None


def _extract_field_numeric(
    rows: list[list[Any]],
    aliases: list[str],
    year_columns: dict[int, int],
    year: int,
    unit_multiplier: float,
    field_name: str | None = None,
):
    year_col = year_columns.get(year)
    blocked_tokens = FIELD_BLOCKED_TOKENS.get(field_name or "", [])
    for row in rows:
        matched = _row_matches_alias(row, aliases)
        if not matched:
            continue
        label_index, _ = matched
        normalized_label = _normalize_text(row[label_index])
        if any(token in normalized_label for token in blocked_tokens):
            continue
        value = _extract_numeric_value(row, label_index, year_col, unit_multiplier)
        if value is not None:
            return value
    return None


def _extract_currency(rows: list[list[Any]]):
    for row in rows[:30]:
        for cell in row[:6]:
            text = str(cell or "").strip().lower()
            if not text:
                continue
            if "idr" in text or "rupiah" in text:
                return "IDR"
            if "usd" in text or "dollar" in text:
                return "USD"
    return None


def _normalize_period(result: dict) -> str:
    raw = str(result.get("Report_Period") or result.get("report_period") or "").strip().lower()
    if "audit" in raw or "annual" in raw or "tahunan" in raw:
        return "AUDIT"
    if any(token in raw for token in ["q1", "tw1", "triwulan i", "triwulan 1"]):
        return "Q1"
    if any(token in raw for token in ["q2", "tw2", "triwulan ii", "triwulan 2"]):
        return "Q2"
    if any(token in raw for token in ["q3", "tw3", "triwulan iii", "triwulan 3"]):
        return "Q3"
    if any(token in raw for token in ["q4", "tw4", "triwulan iv", "triwulan 4"]):
        return "Q4"
    return "AUDIT"


def _period_from_file_name(file_name: str) -> str | None:
    lower_name = (file_name or "").lower()
    if not lower_name:
        return None

    if any(token in lower_name for token in ["tahunan", "annual", "audit"]):
        return "AUDIT"

    # Handle roman numerals in order from longest to shortest to avoid partial matches.
    if re.search(r"(?:^|[-_\s])iv(?:[-_\s.]|$)", lower_name):
        return "Q4"
    if re.search(r"(?:^|[-_\s])iii(?:[-_\s.]|$)", lower_name):
        return "Q3"
    if re.search(r"(?:^|[-_\s])ii(?:[-_\s.]|$)", lower_name):
        return "Q2"
    if re.search(r"(?:^|[-_\s])i(?:[-_\s.]|$)", lower_name):
        return "Q1"

    if any(token in lower_name for token in ["q1", "tw1", "triwulan1", "quarter1"]):
        return "Q1"
    if any(token in lower_name for token in ["q2", "tw2", "triwulan2", "quarter2"]):
        return "Q2"
    if any(token in lower_name for token in ["q3", "tw3", "triwulan3", "quarter3"]):
        return "Q3"
    if any(token in lower_name for token in ["q4", "tw4", "triwulan4", "quarter4"]):
        return "Q4"

    return None


def _resolve_period(result: dict, file_name: str) -> str:
    by_file_name = _period_from_file_name(file_name)
    if by_file_name:
        return by_file_name
    raw = str(result.get("Report_Period") or result.get("report_period") or "").strip().lower()
    if "audit" in raw or "annual" in raw or "tahunan" in raw:
        return "AUDIT"
    if any(token in raw for token in ["q1", "tw1", "triwulan i", "triwulan 1"]):
        return "Q1"
    if any(token in raw for token in ["q2", "tw2", "triwulan ii", "triwulan 2"]):
        return "Q2"
    if any(token in raw for token in ["q3", "tw3", "triwulan iii", "triwulan 3"]):
        return "Q3"
    if any(token in raw for token in ["q4", "tw4", "triwulan iv", "triwulan 4"]):
        return "Q4"
    return "AUDIT"


def _fiscal_quarter(period: str):
    mapping = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
    return mapping.get(period)


def _period_end_date(fiscal_year: int, fiscal_quarter: int | None) -> str:
    if fiscal_quarter == 1:
        return f"{fiscal_year}-03-31"
    if fiscal_quarter == 2:
        return f"{fiscal_year}-06-30"
    if fiscal_quarter == 3:
        return f"{fiscal_year}-09-30"
    return f"{fiscal_year}-12-31"


def _audit_status(period: str) -> str:
    if period == "AUDIT":
        return "AUDITED"
    return "UNAUDITED"


def _extract_sheet_metrics(rows: list[list[Any]], year: int) -> dict[str, Any]:
    year_columns = _find_year_columns(rows, year)
    unit_multiplier = _detect_unit_multiplier(rows)
    metrics: dict[str, Any] = {}

    for field, aliases in NUMERIC_ALIASES.items():
        value = _extract_field_numeric(rows, aliases, year_columns, year, unit_multiplier, field_name=field)
        if value is not None:
            metrics[field] = value
        else:
            metrics[field] = None

    current_revenue = metrics.get("revenue")
    previous_revenue = None
    if (year - 1) in year_columns:
        previous_revenue = _extract_field_numeric(
            rows,
            NUMERIC_ALIASES["revenue"],
            year_columns,
            year - 1,
            unit_multiplier,
            field_name="revenue",
        )

    if current_revenue not in (None, 0) and previous_revenue not in (None, 0):
        metrics["revenueGrowthYoY"] = round((current_revenue - previous_revenue) / abs(previous_revenue), 6)
    else:
        metrics["revenueGrowthYoY"] = None

    # Normalize effective tax rate to ratio when source is percentage.
    metrics["effectiveTaxRate"] = _to_ratio(metrics.get("effectiveTaxRate"))
    metrics["currency"] = _extract_currency(rows)

    # Bank fallback: when revenue is missing/too small, use operating profile approximation.
    if metrics.get("revenue") in (None, 0) or (
        metrics.get("interestIncome") not in (None, 0)
        and metrics.get("revenue") is not None
        and abs(metrics.get("revenue") or 0) < abs(metrics.get("interestIncome") or 0) * 0.2
    ):
        interest_income = metrics.get("interestIncome") or 0
        operating_income = metrics.get("operatingIncome") or 0
        operating_expenses = metrics.get("operatingExpenses") or 0
        estimated_revenue = max(interest_income, operating_income + operating_expenses)
        metrics["revenue"] = estimated_revenue if estimated_revenue > 0 else metrics.get("revenue")

    return metrics


def _extract_balance_metrics(rows: list[list[Any]], year: int) -> dict[str, Any]:
    year_columns = _find_year_columns(rows, year)
    unit_multiplier = _detect_unit_multiplier(rows)
    metrics: dict[str, Any] = {}

    for field, aliases in BALANCE_NUMERIC_ALIASES.items():
        metrics[field] = _extract_field_numeric(rows, aliases, year_columns, year, unit_multiplier, field_name=field)

    metrics["currency"] = _extract_currency(rows)

    # Derived values
    current_assets = metrics.get("totalCurrentAssets")
    current_liabilities = metrics.get("totalCurrentLiabilities")
    cash = metrics.get("cash")
    short_debt = metrics.get("shortTermDebt")
    long_debt = metrics.get("longTermDebt")
    total_equity = metrics.get("totalEquity")

    if isinstance(current_assets, (int, float)) and isinstance(current_liabilities, (int, float)):
        metrics["workingCapital"] = float(current_assets) - float(current_liabilities)
    else:
        metrics["workingCapital"] = None

    debt_total = 0.0
    has_debt = False
    if isinstance(short_debt, (int, float)):
        debt_total += float(short_debt)
        has_debt = True
    if isinstance(long_debt, (int, float)):
        debt_total += float(long_debt)
        has_debt = True
    if has_debt:
        metrics["netDebt"] = debt_total - float(cash or 0.0)
    else:
        metrics["netDebt"] = None

    # Try to derive residual buckets when totals are present.
    known_current_assets = sum(
        float(metrics.get(key) or 0.0)
        for key in ["cash", "shortTermInvestments", "accountsReceivable", "inventory"]
        if isinstance(metrics.get(key), (int, float))
    )
    if metrics.get("otherCurrentAssets") is None and isinstance(current_assets, (int, float)):
        residual = float(current_assets) - known_current_assets
        metrics["otherCurrentAssets"] = residual if abs(residual) > 1e-6 else 0.0

    known_non_current_assets = sum(
        float(metrics.get(key) or 0.0)
        for key in ["propertyPlantEquipment", "intangibleAssets", "goodwill", "longTermInvestments"]
        if isinstance(metrics.get(key), (int, float))
    )
    if metrics.get("otherNonCurrentAssets") is None and isinstance(metrics.get("totalNonCurrentAssets"), (int, float)):
        residual = float(metrics["totalNonCurrentAssets"]) - known_non_current_assets
        metrics["otherNonCurrentAssets"] = residual if abs(residual) > 1e-6 else 0.0

    if metrics.get("totalNonCurrentAssets") is None and isinstance(metrics.get("totalAssets"), (int, float)) and isinstance(current_assets, (int, float)):
        metrics["totalNonCurrentAssets"] = float(metrics["totalAssets"]) - float(current_assets)

    known_current_liabilities = sum(
        float(metrics.get(key) or 0.0)
        for key in ["shortTermDebt", "accountsPayable", "deferredRevenue"]
        if isinstance(metrics.get(key), (int, float))
    )
    if metrics.get("otherCurrentLiabilities") is None and isinstance(current_liabilities, (int, float)):
        residual = float(current_liabilities) - known_current_liabilities
        metrics["otherCurrentLiabilities"] = residual if abs(residual) > 1e-6 else 0.0

    if metrics.get("totalNonCurrentLiabilities") is None and isinstance(metrics.get("totalLiabilities"), (int, float)) and isinstance(current_liabilities, (int, float)):
        metrics["totalNonCurrentLiabilities"] = float(metrics["totalLiabilities"]) - float(current_liabilities)

    known_non_current_liabilities = sum(
        float(metrics.get(key) or 0.0)
        for key in ["longTermDebt", "deferredTaxLiabilities"]
        if isinstance(metrics.get(key), (int, float))
    )
    if metrics.get("otherNonCurrentLiabilities") is None and isinstance(metrics.get("totalNonCurrentLiabilities"), (int, float)):
        residual = float(metrics["totalNonCurrentLiabilities"]) - known_non_current_liabilities
        metrics["otherNonCurrentLiabilities"] = residual if abs(residual) > 1e-6 else 0.0

    known_equity = sum(
        float(metrics.get(key) or 0.0)
        for key in ["commonStock", "additionalPaidInCapital", "retainedEarnings", "treasuryStock", "minorityInterestEquity"]
        if isinstance(metrics.get(key), (int, float))
    )
    if metrics.get("otherEquity") is None and isinstance(total_equity, (int, float)):
        residual = float(total_equity) - known_equity
        metrics["otherEquity"] = residual if abs(residual) > 1e-6 else 0.0

    shares = metrics.get("sharesWeightedAvg")
    if isinstance(total_equity, (int, float)) and isinstance(shares, (int, float)) and shares not in (0, 0.0):
        metrics["bookValuePerShare"] = round(float(total_equity) / float(shares), 4)
    else:
        metrics["bookValuePerShare"] = None

    return metrics


def _sheet_quality_score(metrics: dict[str, Any]) -> float:
    core_fields = ["revenue", "operatingIncome", "pretaxIncome", "netIncome", "interestIncome", "interestExpense"]
    hits = sum(1 for key in core_fields if metrics.get(key) not in (None, 0))

    magnitude = 0.0
    for key in ["revenue", "operatingIncome", "netIncome", "interestIncome"]:
        value = metrics.get(key)
        if isinstance(value, (int, float)):
            magnitude += abs(float(value))

    return hits * 1_000_000_000_000 + magnitude


def _balance_sheet_quality_score(metrics: dict[str, Any]) -> float:
    core_fields = ["totalAssets", "totalLiabilities", "totalEquity", "totalCurrentAssets", "totalCurrentLiabilities"]
    hits = sum(1 for key in core_fields if metrics.get(key) not in (None, 0))
    magnitude = 0.0
    for key in ["totalAssets", "totalLiabilities", "totalEquity"]:
        value = metrics.get(key)
        if isinstance(value, (int, float)):
            magnitude += abs(float(value))
    return hits * 1_000_000_000_000 + magnitude


def _parse_workbook(content: bytes, year: int) -> tuple[dict[str, Any], dict[str, Any]]:
    wb = load_workbook(filename=BytesIO(content), read_only=True, data_only=True)
    income_sheets = sorted(wb.worksheets, key=lambda s: _sheet_score(s.title), reverse=True)
    balance_sheets = sorted(wb.worksheets, key=lambda s: _balance_sheet_score(s.title), reverse=True)

    merged_income = {field: None for field in NUMERIC_ALIASES.keys()}
    merged_income["revenueGrowthYoY"] = None
    merged_income["currency"] = None
    merged_balance = {field: None for field in BALANCE_NUMERIC_ALIASES.keys()}
    merged_balance["currency"] = None
    merged_balance["bookValuePerShare"] = None
    merged_balance["netDebt"] = None
    merged_balance["workingCapital"] = None

    income_results: list[dict[str, Any]] = []
    balance_results: list[dict[str, Any]] = []

    for sheet in income_sheets:
        rows = _normalize_rows(sheet)
        if not rows:
            continue

        sheet_metrics = _extract_sheet_metrics(rows, year)
        income_results.append(sheet_metrics)

        for key, value in sheet_metrics.items():
            current = merged_income.get(key)
            if current is None and value is not None:
                merged_income[key] = value

    for sheet in balance_sheets:
        rows = _normalize_rows(sheet)
        if not rows:
            continue

        sheet_metrics = _extract_balance_metrics(rows, year)
        balance_results.append(sheet_metrics)

        for key, value in sheet_metrics.items():
            current = merged_balance.get(key)
            if current is None and value is not None:
                merged_balance[key] = value

    if income_results:
        best = max(income_results, key=_sheet_quality_score)
        for key, value in best.items():
            if value is not None:
                merged_income[key] = value

    if balance_results:
        best = max(balance_results, key=_balance_sheet_quality_score)
        for key, value in best.items():
            if value is not None:
                merged_balance[key] = value

    # Use shares from income side when available for BVPS derivation.
    shares = merged_income.get("sharesWeightedAvg")
    if (
        merged_balance.get("bookValuePerShare") is None
        and isinstance(merged_balance.get("totalEquity"), (int, float))
        and isinstance(shares, (int, float))
        and shares not in (0, 0.0)
    ):
        merged_balance["bookValuePerShare"] = round(float(merged_balance["totalEquity"]) / float(shares), 4)

    wb.close()
    return merged_income, merged_balance


def _build_statement_item(result: dict, parsed: dict, fallback_year: int) -> dict:
    fiscal_year = int(result.get("Report_Year") or result.get("report_year") or fallback_year)
    period = _normalize_period(result)
    quarter = _fiscal_quarter(period)

    return {
        "period": period,
        "fiscalYear": fiscal_year,
        "fiscalQuarter": quarter,
        "periodEndDate": _period_end_date(fiscal_year, quarter),
        "currency": parsed.get("currency") or "IDR",
        "auditStatus": _audit_status(period),
        "revenue": parsed.get("revenue"),
        "revenueGrowthYoY": parsed.get("revenueGrowthYoY"),
        "cogs": parsed.get("cogs"),
        "grossProfit": parsed.get("grossProfit"),
        "operatingExpenses": parsed.get("operatingExpenses"),
        "sellingExpenses": parsed.get("sellingExpenses"),
        "generalAdminExpenses": parsed.get("generalAdminExpenses"),
        "rdExpenses": parsed.get("rdExpenses"),
        "depreciationAmort": parsed.get("depreciationAmort"),
        "ebit": parsed.get("ebit"),
        "ebitda": parsed.get("ebitda"),
        "operatingIncome": parsed.get("operatingIncome"),
        "interestExpense": parsed.get("interestExpense"),
        "interestIncome": parsed.get("interestIncome"),
        "otherNonOperatingIncome": parsed.get("otherNonOperatingIncome"),
        "pretaxIncome": parsed.get("pretaxIncome"),
        "incomeTaxExpense": parsed.get("incomeTaxExpense"),
        "effectiveTaxRate": parsed.get("effectiveTaxRate"),
        "netIncome": parsed.get("netIncome"),
        "netIncomeAttributable": parsed.get("netIncomeAttributable"),
        "minorityInterest": parsed.get("minorityInterest"),
        "eps": parsed.get("eps"),
        "epsDiluted": parsed.get("epsDiluted"),
        "sharesWeightedAvg": parsed.get("sharesWeightedAvg"),
    }


def _build_balance_sheet_item(result: dict, parsed: dict, fallback_year: int) -> dict:
    fiscal_year = int(result.get("Report_Year") or result.get("report_year") or fallback_year)
    period = _normalize_period(result)
    quarter = _fiscal_quarter(period)

    return {
        "period": period,
        "fiscalYear": fiscal_year,
        "fiscalQuarter": quarter,
        "periodEndDate": _period_end_date(fiscal_year, quarter),
        "currency": parsed.get("currency") or "IDR",
        "auditStatus": _audit_status(period),
        "cash": parsed.get("cash"),
        "shortTermInvestments": parsed.get("shortTermInvestments"),
        "accountsReceivable": parsed.get("accountsReceivable"),
        "inventory": parsed.get("inventory"),
        "otherCurrentAssets": parsed.get("otherCurrentAssets"),
        "totalCurrentAssets": parsed.get("totalCurrentAssets"),
        "propertyPlantEquipment": parsed.get("propertyPlantEquipment"),
        "intangibleAssets": parsed.get("intangibleAssets"),
        "goodwill": parsed.get("goodwill"),
        "longTermInvestments": parsed.get("longTermInvestments"),
        "otherNonCurrentAssets": parsed.get("otherNonCurrentAssets"),
        "totalNonCurrentAssets": parsed.get("totalNonCurrentAssets"),
        "totalAssets": parsed.get("totalAssets"),
        "shortTermDebt": parsed.get("shortTermDebt"),
        "accountsPayable": parsed.get("accountsPayable"),
        "deferredRevenue": parsed.get("deferredRevenue"),
        "otherCurrentLiabilities": parsed.get("otherCurrentLiabilities"),
        "totalCurrentLiabilities": parsed.get("totalCurrentLiabilities"),
        "longTermDebt": parsed.get("longTermDebt"),
        "deferredTaxLiabilities": parsed.get("deferredTaxLiabilities"),
        "otherNonCurrentLiabilities": parsed.get("otherNonCurrentLiabilities"),
        "totalNonCurrentLiabilities": parsed.get("totalNonCurrentLiabilities"),
        "totalLiabilities": parsed.get("totalLiabilities"),
        "commonStock": parsed.get("commonStock"),
        "additionalPaidInCapital": parsed.get("additionalPaidInCapital"),
        "retainedEarnings": parsed.get("retainedEarnings"),
        "treasuryStock": parsed.get("treasuryStock"),
        "otherEquity": parsed.get("otherEquity"),
        "minorityInterestEquity": parsed.get("minorityInterestEquity"),
        "totalEquity": parsed.get("totalEquity"),
        "bookValuePerShare": parsed.get("bookValuePerShare"),
        "netDebt": parsed.get("netDebt"),
        "workingCapital": parsed.get("workingCapital"),
    }


def _normalize_monetary_scale(item: dict) -> dict:
    monetary_fields = [
        "revenue",
        "cogs",
        "grossProfit",
        "operatingExpenses",
        "sellingExpenses",
        "generalAdminExpenses",
        "rdExpenses",
        "depreciationAmort",
        "ebit",
        "ebitda",
        "operatingIncome",
        "interestExpense",
        "interestIncome",
        "otherNonOperatingIncome",
        "pretaxIncome",
        "incomeTaxExpense",
        "netIncome",
        "netIncomeAttributable",
        "minorityInterest",
    ]

    numeric_values = [
        abs(float(item.get(field)))
        for field in monetary_fields
        if isinstance(item.get(field), (int, float)) and item.get(field) not in (None, 0)
    ]
    if not numeric_values:
        return item

    max_value = max(numeric_values)
    revenue_value = item.get("revenue") if isinstance(item.get("revenue"), (int, float)) else None

    # Heuristic: values are likely reported in millions if all monetary figures are
    # relatively small for IDR but still clearly non-trivial financial statement numbers.
    looks_like_millions = (
        str(item.get("currency") or "").upper() == "IDR"
        and max_value < 1_000_000_000
        and (
            max_value >= 1_000_000
            or (revenue_value is not None and 10_000 <= abs(float(revenue_value)) < 1_000_000_000)
        )
    )
    if not looks_like_millions:
        return item

    for field in monetary_fields:
        value = item.get(field)
        if isinstance(value, (int, float)):
            item[field] = float(value) * 1_000_000

    return item


def _apply_bank_derivations(item: dict) -> dict:
    interest_income = item.get("interestIncome") if isinstance(item.get("interestIncome"), (int, float)) else None
    interest_expense = item.get("interestExpense") if isinstance(item.get("interestExpense"), (int, float)) else None
    revenue = item.get("revenue") if isinstance(item.get("revenue"), (int, float)) else None
    operating_income = item.get("operatingIncome") if isinstance(item.get("operatingIncome"), (int, float)) else None
    other_non_op = item.get("otherNonOperatingIncome")
    if not isinstance(other_non_op, (int, float)):
        other_non_op = 0.0
        item["otherNonOperatingIncome"] = 0.0

    if item.get("cogs") is None and interest_expense is not None:
        item["cogs"] = float(interest_expense)

    if item.get("grossProfit") is None and isinstance(item.get("revenue"), (int, float)) and isinstance(item.get("cogs"), (int, float)):
        item["grossProfit"] = float(item["revenue"]) - float(item["cogs"])

    if item.get("ebit") is None and operating_income is not None:
        item["ebit"] = float(operating_income)

    if item.get("pretaxIncome") is None and operating_income is not None:
        item["pretaxIncome"] = float(operating_income) + float(other_non_op)

    if item.get("incomeTaxExpense") is None and isinstance(item.get("pretaxIncome"), (int, float)) and isinstance(item.get("netIncome"), (int, float)):
        item["incomeTaxExpense"] = float(item["pretaxIncome"]) - float(item["netIncome"])

    if isinstance(item.get("incomeTaxExpense"), (int, float)) and isinstance(item.get("pretaxIncome"), (int, float)):
        if abs(float(item["incomeTaxExpense"]) - float(item["pretaxIncome"])) < 1e-9:
            # Clearly misclassified from "profit before tax" row.
            item["incomeTaxExpense"] = None

    if item.get("netIncome") is None and isinstance(item.get("pretaxIncome"), (int, float)) and isinstance(item.get("incomeTaxExpense"), (int, float)):
        item["netIncome"] = float(item["pretaxIncome"]) - float(item["incomeTaxExpense"])

    if item.get("netIncomeAttributable") is None and isinstance(item.get("netIncome"), (int, float)) and isinstance(item.get("minorityInterest"), (int, float)):
        item["netIncomeAttributable"] = float(item["netIncome"]) - float(item["minorityInterest"])

    if item.get("netIncome") is None and isinstance(item.get("netIncomeAttributable"), (int, float)) and isinstance(item.get("minorityInterest"), (int, float)):
        item["netIncome"] = float(item["netIncomeAttributable"]) + float(item["minorityInterest"])

    if isinstance(item.get("pretaxIncome"), (int, float)) and isinstance(item.get("netIncome"), (int, float)):
        pretax = float(item["pretaxIncome"])
        net = float(item["netIncome"])
        if pretax > 0 and net > pretax * 1.2:
            item["netIncome"] = None
            if isinstance(item.get("netIncomeAttributable"), (int, float)) and isinstance(item.get("minorityInterest"), (int, float)):
                item["netIncome"] = float(item["netIncomeAttributable"]) + float(item["minorityInterest"])

    if item.get("effectiveTaxRate") is None and isinstance(item.get("incomeTaxExpense"), (int, float)) and isinstance(item.get("pretaxIncome"), (int, float)) and item["pretaxIncome"] not in (0, 0.0):
        item["effectiveTaxRate"] = round(abs(float(item["incomeTaxExpense"])) / abs(float(item["pretaxIncome"])), 6)

    if isinstance(item.get("incomeTaxExpense"), (int, float)) and isinstance(item.get("pretaxIncome"), (int, float)):
        if abs(float(item["incomeTaxExpense"])) > abs(float(item["pretaxIncome"])) * 0.9:
            item["incomeTaxExpense"] = None
            item["effectiveTaxRate"] = None

    if (
        isinstance(item.get("pretaxIncome"), (int, float))
        and isinstance(item.get("netIncome"), (int, float))
        and abs(float(item["pretaxIncome"]) - float(item["netIncome"])) < 1e-9
        and item.get("incomeTaxExpense") in (None, 0, 0.0)
    ):
        assumed_rate = 0.2
        pretax = float(item["pretaxIncome"])
        item["incomeTaxExpense"] = round(pretax * assumed_rate, 2)
        item["netIncome"] = round(pretax - item["incomeTaxExpense"], 2)
        item["effectiveTaxRate"] = assumed_rate
        if isinstance(item.get("minorityInterest"), (int, float)):
            item["netIncomeAttributable"] = round(float(item["netIncome"]) - float(item["minorityInterest"]), 2)

    if revenue in (None, 0) and interest_income is not None:
        item["revenue"] = float(interest_income)

    return item


def scrape_financial_statement(symbol: str, year: int) -> dict:
    results = fetch_financial_report_results(symbol, year)
    spreadsheets = _collect_spreadsheet_attachments(results)

    income_items: list[dict] = []
    balance_items: list[dict] = []
    seen_income_period_year: set[tuple[str, int]] = set()
    seen_balance_period_year: set[tuple[str, int]] = set()

    for result, attachment in spreadsheets:
        file_name = str(attachment.get("File_Name") or attachment.get("file_name") or "")
        file_url = _attachment_url(attachment)
        if not file_url:
            continue

        try:
            content = _download_file(file_url)
            parsed_income, parsed_balance = _parse_workbook(content, year)

            income_item = _build_statement_item(result, parsed_income, fallback_year=year)
            income_item["period"] = _resolve_period(result, file_name)
            income_item["fiscalQuarter"] = _fiscal_quarter(income_item["period"])
            income_item["periodEndDate"] = _period_end_date(int(income_item["fiscalYear"]), income_item["fiscalQuarter"])
            income_item["auditStatus"] = _audit_status(income_item["period"])

            balance_item = _build_balance_sheet_item(result, parsed_balance, fallback_year=year)
            balance_item["period"] = _resolve_period(result, file_name)
            balance_item["fiscalQuarter"] = _fiscal_quarter(balance_item["period"])
            balance_item["periodEndDate"] = _period_end_date(int(balance_item["fiscalYear"]), balance_item["fiscalQuarter"])
            balance_item["auditStatus"] = _audit_status(balance_item["period"])
        except Exception:
            continue

        income_dedup_key = (str(income_item.get("period") or ""), int(income_item.get("fiscalYear") or year))
        if income_dedup_key not in seen_income_period_year:
            income_numeric_hits = sum(
                1 for key in ["revenue", "operatingIncome", "pretaxIncome", "netIncome", "ebit"] if income_item.get(key) not in (None, 0)
            )
            if income_numeric_hits > 0:
                seen_income_period_year.add(income_dedup_key)
                income_item = _normalize_monetary_scale(income_item)
                income_item = _apply_bank_derivations(income_item)
                income_items.append(income_item)

        balance_dedup_key = (str(balance_item.get("period") or ""), int(balance_item.get("fiscalYear") or year))
        if balance_dedup_key not in seen_balance_period_year:
            balance_numeric_hits = sum(
                1
                for key in ["totalAssets", "totalLiabilities", "totalEquity", "totalCurrentAssets", "totalCurrentLiabilities"]
                if balance_item.get(key) not in (None, 0)
            )
            if balance_numeric_hits > 0:
                seen_balance_period_year.add(balance_dedup_key)
                balance_item = _normalize_monetary_scale(balance_item)
                balance_items.append(balance_item)

    income_items.sort(key=lambda row: (int(row.get("fiscalYear") or 0), int(row.get("fiscalQuarter") or 99)))
    balance_items.sort(key=lambda row: (int(row.get("fiscalYear") or 0), int(row.get("fiscalQuarter") or 99)))
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
    }


def scrape_income_statement(symbol: str, year: int) -> dict:
    # Backward-compatible alias for internal imports that may still call old name.
    return scrape_financial_statement(symbol, year)
