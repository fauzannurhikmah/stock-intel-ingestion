from __future__ import annotations

from typing import List, Tuple, Dict

from app.scrapers.common import BASE_URL, _get
from app.scrapers.fs_utilities.fs_utils import _normalize_file_extension, _score_attachment, SUPPORTED_EXTENSIONS, MAX_ATTACHMENTS_TO_PARSE


def fetch_financial_report_results(symbol: str, year: int) -> list[dict]:
    url = f"{BASE_URL}/primary/ListedCompany/GetFinancialReport"
    params = {
        "ReportType": "rdf",
        "KodeEmiten": symbol.upper(),
        "Year": str(year),
        "SortColumn": "KodeEmiten",
        "SortOrder": "asc",
        "EmitenType": None,
        "Periode": "*",
        "indexfrom": 0,
        "pagesize": 0,
    }
    response_data = _get(url, params)
    results = response_data.get("Results") or response_data.get("results") or []
    return results if isinstance(results, list) else []


def _collect_pdf_attachments(results: list[dict]) -> List[Tuple[Dict, Dict]]:
    collected: List[Tuple[Dict, Dict]] = []
    REPORT_KEYWORDS = [
        "laporankeuangan",
        "laporan keuangan",
        "lapkeu",
        "lap keu",
        "financialstatement",
        "financial statement",
        "laporan keuangan konsolidasian",
    ]

    def _matches_report_keyword(attachment: dict) -> bool:
        name = str(attachment.get("File_Name") or attachment.get("file_name") or "").lower()
        path = str(attachment.get("File_Path") or attachment.get("file_path") or "").lower()
        combined = f"{name} {path}"
        return any(kw in combined for kw in REPORT_KEYWORDS)

    for result in results:
        attachments = result.get("Attachments") or result.get("attachments") or []
        if not isinstance(attachments, list):
            continue

        pdfs: list[dict] = []
        for attachment in attachments:
            if not isinstance(attachment, dict):
                continue
            file_name = str(attachment.get("File_Name") or attachment.get("file_name") or "")
            file_type = str(attachment.get("File_Type") or attachment.get("file_type") or "")
            ext = _normalize_file_extension(file_name, file_type)
            if ext == ".pdf":
                pdfs.append(attachment)

        if not pdfs:
            continue

        preferred = [a for a in pdfs if _matches_report_keyword(a)]
        chosen = preferred if preferred else pdfs

        for attachment in chosen:
            collected.append((result, attachment))

    collected.sort(key=lambda item: _score_attachment(item[1]), reverse=True)
    return collected[:MAX_ATTACHMENTS_TO_PARSE]


# Backwards-compatible alias
_collect_spreadsheet_attachments = _collect_pdf_attachments
