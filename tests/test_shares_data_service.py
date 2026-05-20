from app.scraper import _parse_shares_report_text, _select_monthly_announcements


def test_select_monthly_announcements_prefers_koreksi():
    replies = [
        {
            "pengumuman": {
                "JudulPengumuman": "Laporan Bulanan Registrasi Pemegang Efek",
                "TglPengumuman": "2026-05-08T18:12:00",
            },
            "attachments": [{"FullSavePath": "https://example.com/may.pdf"}],
        },
        {
            "pengumuman": {
                "JudulPengumuman": "Laporan Bulanan Registrasi Pemegang Efek (KOREKSI)",
                "TglPengumuman": "2026-05-30T18:12:00",
            },
            "attachments": [{"FullSavePath": "https://example.com/may-correction.pdf"}],
        },
        {
            "pengumuman": {
                "JudulPengumuman": "Laporan Bulanan Registrasi Pemegang Efek",
                "TglPengumuman": "2026-04-30T18:12:00",
            },
            "attachments": [{"FullSavePath": "https://example.com/apr.pdf"}],
        },
    ]

    selected = _select_monthly_announcements(replies)

    assert len(selected) == 2
    assert selected[-1]["title"].endswith("(KOREKSI)")


def test_parse_shares_report_text_extracts_metrics():
    text = """
    Laporan Bulanan Registrasi Pemegang Efek
    Total Saham Beredar 3.500.000.000
    Free Float 1.250.000.000
    Kepemilikan Institusional 1.900.000.000
    Kepemilikan Insider 350.000.000
    """

    parsed = _parse_shares_report_text(text, announcement_date="2026-05-30")

    assert parsed["date"] == "2026-05-30"
    assert parsed["sharesOutstanding"] == 3500000000
    assert parsed["sharesFloat"] == 1250000000
    assert parsed["sharesInstitutional"] == 1900000000
    assert parsed["sharesInsider"] == 350000000