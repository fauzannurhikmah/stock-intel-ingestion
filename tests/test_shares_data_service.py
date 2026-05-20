from app.scraper import (
    _fill_missing_share_metrics,
    _parse_shares_report_text,
    _select_monthly_announcements,
)


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


def test_parse_shares_report_text_extracts_idx_lbre_table_metrics():
    text = """
    Pemegang saham 5% (lima persen) atau lebih
    Total
    123.275.050.
    000
    100% 123.275.050.
    000
    100%
    Jumlah Saham Non Warkat berdasarkan Tipe Investor
    Tipe dan Klasifikasi Investor Jumlah Saham
    Afiliasi dari pengendali dan pemilik manfaat
    Direksi dan Dewan Komisaris
    Karyawan (hanya saham yang dibatasi pengalihan kepemilikannya)
    Pengendali
    Saham Tresuri
    3.082.137.500
    80.375.642
    0
    67.729.950.000
    401.316.300
    Jumlah Saham Scripless berdasarkan Tipe dan Klasifikasi Investor dari KSEI
    Securities Company (SC)
    5.660.775.303
    330.790.081
    Pension Funds (PF)
    310.478.277
    18.920.621.304
    Financial Institutional (IB)
    1.711.122.286
    Partnership
    Educational Institution
    Mutual Funds (MF)
    213.056.686
    Foundation (FD)
    Individual (ID)
    341.660.165
    10.778.610.902
    Informasi Saham Free Float
    Jumlah saham Free Float
    Jumlah saham tercatat di Bursa per akhir bulan
    % Saham Free Float
    52.453.691.120 52.453.691.120
    80.375.642 80.375.642
    1.977.500
    0
    """

    parsed = _parse_shares_report_text(text, announcement_date="2026-05-08")

    assert parsed["date"] == "2026-05-08"
    assert parsed["sharesOutstanding"] == 123275050000
    assert parsed["sharesFloat"] == 52453691120
    assert parsed["sharesInstitutional"] == 1711122286
    assert parsed["sharesInsider"] == 80375642


def test_parse_shares_report_text_extracts_insider_from_free_float_table():
    text = """
    Informasi Saham Free Float
    Keterangan Bulan Sebelumnya Bulan Ini
    Jumlah saham scripless dimiliki oleh pemegang saham kurang
    dari 5%
    Jumlah saham scripless dimiliki oleh Direksi dan Dewan
    Komisaris kurang dari 5%
    Jumlah saham scripless dimiliki oleh Pengendali kurang dari 5%
    Jumlah saham scripless dimiliki oleh Afiliasi dari Pengendali
    kurang dari 5%
    Jumlah treasury stock scripless kurang dari 5%
    Jumlah saham portofolio investasi dengan penerima manfaat
    investor publik yang telah disetujui Bursa untuk diperhitungkan
    sebagai Free Float
    Jumlah saham Free Float
    Jumlah saham tercatat di Bursa per akhir bulan
    % Saham Free Float
    52,453,696,120 52,453,696,120
    80,375,642 80,375,642
    0 0
    0 0
    0 0
    123,275,050,000 123,275,050,000
    """

    parsed = _parse_shares_report_text(text, announcement_date="2023-07-05")

    assert parsed["sharesFloat"] == 52453696120
    assert parsed["sharesInsider"] == 80375642


def test_parse_shares_report_text_extracts_legacy_free_float_and_insider():
    text = """
    Total
    123.275.050.
    000
    100% 123.275.050.
    000
    100%
    Laporan Kepemilikan Saham Oleh Direksi dan Komisaris
    Nama Jabatan Alamat
    Jumlah Saham Bulan Sebelumnya
    Persen Saham Bulan Sebelumnya
    Jumlah Saham Bulan Ini
    Persen Saham Bulan Ini
    Jahja Setiaatmadja Direksi Menara BCA Jakarta 10310
    40.818.853 0,03% 40.818.853 0,03%
    Armand Wahyudi Direksi Menara BCA Jakarta 10310
    4.256.065 0% 4.256.065 0%
    Informasi Saham Free Float
    Jumlah Saham Bulan Sebelumnya
    Persen Saham Bulan Sebelumnya
    Jumlah Saham Bulan Ini
    Persen Saham Bulan Ini
    Jumlah Perubahan Saham
    Persen Perubahan
    52.265.846.217 52.266.846.217 1.000.00042,4% 42,4% 0%
    """

    parsed = _parse_shares_report_text(text, announcement_date="2023-06-07")

    assert parsed["sharesOutstanding"] == 123275050000
    assert parsed["sharesFloat"] == 52266846217
    assert parsed["sharesInsider"] == 45074918


def test_fill_missing_share_metrics_carries_core_values_and_zeroes_institutional():
    items = [
        {
            "date": "2023-09-06",
            "sharesOutstanding": 123275050000,
            "sharesFloat": 52453696120,
            "sharesInsider": 80375642,
            "sharesInstitutional": None,
        },
        {
            "date": "2023-10-06",
            "sharesOutstanding": None,
            "sharesFloat": None,
            "sharesInsider": None,
            "sharesInstitutional": None,
        },
        {
            "date": "2023-11-06",
            "sharesOutstanding": 123275050000,
            "sharesFloat": 52453696120,
            "sharesInsider": 80375642,
            "sharesInstitutional": None,
        },
    ]

    filled = _fill_missing_share_metrics(items)

    assert filled[1]["sharesOutstanding"] == 123275050000
    assert filled[1]["sharesFloat"] == 52453696120
    assert filled[1]["sharesInsider"] == 80375642
    assert all(item["sharesInstitutional"] == 0 for item in filled)
