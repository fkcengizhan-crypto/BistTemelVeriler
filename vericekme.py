# -*- coding: utf-8 -*-
"""
BIST Veri Çekici — Tek Dosya
=============================
TradingView canlı verisiyle BIST hisse verilerini çeker ve GÜNLÜK ARŞİV
Excel dosyası üretir:  Archive/Hisse_Verileri_YYYYMMDD.xlsx
(Aynı gün tekrar çalıştırılırsa yeni dosya oluşturulmaz, üzerine yazılır.)
Puanlama/skorlama YOKTUR; yalnızca ham + normalize edilmiş veri listelenir.

Kullanım:
    python vericekme.py
    python vericekme.py --hisse THYAO
    python vericekme.py --sektor Banka
    python vericekme.py --top 100
    python vericekme.py --offline
    python vericekme.py --cikti C:\\Raporlar

Çıktı varsayılan olarak bu script'in yanındaki "Archive" klasörüne yazılır.

Gereksinim: Python 3.9+, openpyxl, pandas
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import unicodedata
import re
import urllib.request
import warnings
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
warnings.filterwarnings("ignore")

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows
except ImportError:
    print("openpyxl gerekli: pip install openpyxl", file=sys.stderr)
    raise

import pandas as pd


# ---------------------------------------------------------------------------
# Sabitler
# ---------------------------------------------------------------------------

# Proje kökü ve günlük arşiv klasörü (script'in bulunduğu dizin esas alınır,
# böylece proje nereden çalıştırılırsa çalıştırılsın doğru yere yazar).
BASE_DIR = Path(__file__).resolve().parent
ARCHIVE_DIR = BASE_DIR / "Archive"
DEFAULT_OFFLINE = BASE_DIR / "Bist_Hisse_Tum_Veriler.xlsx"

MARKET = "turkey"
SCANNER_URL = "https://scanner.tradingview.com/{market}/scan"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Origin": "https://tr.tradingview.com",
    "Referer": "https://tr.tradingview.com/",
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}

# Canlı çekilecek TradingView sütunları
TV_COLUMNS = [
    "name", "description", "sector", "industry", "type",
    "close", "volume", "change",
    "market_cap_basic", "price_52_week_high", "price_52_week_low",
    "price_earnings_ttm", "price_book_fq",
    "dividends_yield", "beta_1_year",
    "enterprise_value_ebitda_ttm",
    "return_on_equity", "return_on_assets", "return_on_invested_capital",
    "gross_margin", "net_margin", "ebitda_margin_ttm",
    "current_ratio", "quick_ratio", "debt_to_equity",
    "net_debt_to_ebitda_fy",
    "altman_z_score_ttm", "piotroski_f_score_ttm",
    "total_revenue_yoy_growth_ttm", "net_income_yoy_growth_ttm",
    "ebitda_yoy_growth_ttm", "free_cash_flow_yoy_growth_ttm",
    "total_revenue_cagr_5y", "net_income_cagr_5y",
    "Perf.W", "Perf.1M", "Perf.3M", "Perf.6M",
    "Perf.Y", "Perf.3Y", "Perf.5Y", "Perf.YTD",
    "Recommend.All", "Recommend.MA", "Recommend.Other",
    "Volatility.M",
    "relative_volume_10d_calc", "average_volume_30d_calc",
    "float_shares_percent_current", "number_of_employees",
    "earnings_release_date", "free_cash_flow_ttm",
]

# TV anahtarı -> dahili anahtar
TV_MAP = {
    "symbol": None,
    "name": "name",
    "company": "description",
    "sector": "sector",
    "industry": "industry",
    "asset_type": "type",
    "close": "close",
    "volume": "volume",
    "change_pct": "change",
    "market_cap": "market_cap_basic",
    "high_52w": "price_52_week_high",
    "low_52w": "price_52_week_low",
    "pe": "price_earnings_ttm",
    "pb": "price_book_fq",
    "div_yield": "dividends_yield",
    "beta": "beta_1_year",
    "ev_ebitda": "enterprise_value_ebitda_ttm",
    "roe": "return_on_equity",
    "roa": "return_on_assets",
    "roic": "return_on_invested_capital",
    "gross_margin": "gross_margin",
    "net_margin": "net_margin",
    "ebitda_margin": "ebitda_margin_ttm",
    "current_ratio": "current_ratio",
    "quick_ratio": "quick_ratio",
    "d_e": "debt_to_equity",
    "net_debt_ebitda": "net_debt_to_ebitda_fy",
    "altman_z": "altman_z_score_ttm",
    "piotroski_f": "piotroski_f_score_ttm",
    "rev_growth": "total_revenue_yoy_growth_ttm",
    "ni_growth": "net_income_yoy_growth_ttm",
    "ebitda_growth": "ebitda_yoy_growth_ttm",
    "fcf_growth": "free_cash_flow_yoy_growth_ttm",
    "rev_cagr5": "total_revenue_cagr_5y",
    "ni_cagr5": "net_income_cagr_5y",
    "perf_1w": "Perf.W",
    "perf_1m": "Perf.1M",
    "perf_3m": "Perf.3M",
    "perf_6m": "Perf.6M",
    "perf_1y": "Perf.Y",
    "perf_3y": "Perf.3Y",
    "perf_5y": "Perf.5Y",
    "perf_ytd": "Perf.YTD",
    "rec_all": "Recommend.All",
    "rec_ma": "Recommend.MA",
    "rec_other": "Recommend.Other",
    "volatility_m": "Volatility.M",
    "rel_volume": "relative_volume_10d_calc",
    "avg_volume": "average_volume_30d_calc",
    "float_pct": "float_shares_percent_current",
    "employees": "number_of_employees",
    "earnings_date": "earnings_release_date",
    "fcf": "free_cash_flow_ttm",
}

# Offline Excel (Türkçe kolon) -> dahili anahtar
TR_MAP = {
    "BIST Kodu": "symbol",
    "Hisse Adı": "name",
    "Şirket Tam Adı": "company",
    "Sektör": "sector",
    "Alt Sektör": "industry",
    "Tür": "asset_type",
    "Son Fiyat": "close",
    "Hacim": "volume",
    "Günlük Değişim (%)": "change_pct",
    "Piyasa Değeri": "market_cap",
    "52 Haftalık En Yüksek": "high_52w",
    "52 Haftalık En Düşük": "low_52w",
    "F/K (TTM)": "pe",
    "F/D (Çeyreklik)": "pb",
    "Temettü Verimi (%)": "div_yield",
    "Beta (1 Yıllık)": "beta",
    "Firma Değeri / FAVÖK (TTM)": "ev_ebitda",
    "Özkaynak Getirisi (ROE %)": "roe",
    "Varlık Getirisi (ROA %)": "roa",
    "Yatırılmış Sermaye Getirisi (ROIC %)": "roic",
    "Brüt Marj (TTM)": "gross_margin",
    "Net Kar Marjı (%)": "net_margin",
    "Ebitda Marj (TTM)": "ebitda_margin",
    "Current Oran (Yıllık)": "current_ratio",
    "Quick Oran (Yıllık)": "quick_ratio",
    "Borç / Özkaynak": "d_e",
    "Net Borç / Ebitda (Yıllık)": "net_debt_ebitda",
    "Altman z Skor (TTM)": "altman_z",
    "Piotroski f Skor (TTM)": "piotroski_f",
    "Toplam Gelir Büyüme (TTM": "rev_growth",
    "Net Kâr Büyüme (TTM": "ni_growth",
    "Ebitda Büyüyme (TTM, YOY)": "ebitda_growth",
    "Serbest Nakit Akış Büyüyme (TTM, YOY)": "fcf_growth",
    "Toplam Gelir 5 Yıllık Bileşik Büyüme (%)": "rev_cagr5",
    "Net Kâr 5 Yıllık Bileşik Büyüme (%)": "ni_cagr5",
    "Haftalık Performans": "perf_1w",
    "1 Aylık Performans": "perf_1m",
    "3 Aylık Performans": "perf_3m",
    "6 Aylık Performans": "perf_6m",
    "1 Yıllık Performans": "perf_1y",
    "Performans (3 Yıl)": "perf_3y",
    "5 Yıllık Performans": "perf_5y",
    "Performans (Yıl Başı)": "perf_ytd",
    "Teknik Tavsiye (Genel)": "rec_all",
    "Aylık Volatilite (%)": "volatility_m",
    "Göreli Hacim (10 Günlük)": "rel_volume",
    "Ortalama Hacim (30 Günlük)": "avg_volume",
    "Halka Açık Hisse Yüzde Current": "float_pct",
    "Çalışan Sayısı": "employees",
    "Bir Sonraki Finansal Rapor Tarihi": "earnings_date",
    "Serbest Nakit Akış (TTM)": "fcf",
}

SECTOR_MAP = {
    "Finance": "Finans",
    "Process Industries": "İşlem Sanayii",
    "Producer Manufacturing": "Üretim Sanayi",
    "Consumer Non-Durables": "Tüketim Ürünleri (Dayanıksız)",
    "Non-Energy Minerals": "Enerji Dışı Madencilik",
    "Utilities": "Kamu Hizmetleri",
    "Consumer Services": "Tüketim Hizmetleri",
    "Consumer Durables": "Tüketim Ürünleri (Dayanıklı)",
    "Miscellaneous": "Diğer",
    "Technology Services": "Teknoloji Hizmetleri",
    "Distribution Services": "Dağıtım Hizmetleri",
    "Retail Trade": "Perakende",
    "Industrial Services": "Endüstriyel Hizmetler",
    "Electronic Technology": "Elektronik Teknoloji",
    "Commercial Services": "Ticari Hizmetler",
    "Transportation": "Ulaşım",
    "Health Technology": "Sağlık Teknolojisi",
    "Health Services": "Sağlık Hizmetleri",
    "Energy Minerals": "Enerji Madenciliği",
    "Communications": "İletişim",
    "Real Estate Investment Trusts": "Gayrimenkul Yatırım Ortaklıkları",
    "Real Estate Development": "Gayrimenkul Geliştirme",
    "Construction Materials": "İnşaat Malzemeleri",
    "Investment Managers": "Yatırım Yöneticileri",
    "Finance/Rental/Leasing": "Finans / Kiralama / Leasing",
    "Investment Banks/Brokers": "Yatırım Bankaları / Aracı Kurumlar",
    "Financial Conglomerates": "Finans Holdingleri",
    "Major Banks": "Büyük Bankalar",
    "Regional Banks": "Bölgesel Bankalar",
    "Airlines": "Havayolları",
    "Oil Refining/Marketing": "Petrol Rafinerisi / Pazarlama",
    "Steel": "Çelik",
    "Wireless Telecommunications": "Kablosuz Telekomünikasyon",
    "Major Telecommunications": "Büyük Telekom",
    "Food Retail": "Gıda Perakende",
    "Packaged Software": "Paket Yazılım",
    "Aerospace & Defense": "Havacılık ve Savunma",
    "Property/Casualty Insurance": "Gayrimenkul / Kaza Sigortası",
    "Life/Health Insurance": "Hayat / Sağlık Sigortası",
    "Multi-Line Insurance": "Çok Branşlı Sigorta",
    "Integrated Oil": "Entegre Petrol",
    "Alternative Power Generation": "Yenilenebilir Enerji Üretimi",
}

COLORS = {
    "header_bg": "1F4E79",
    "header_font": "FFFFFF",
    "subheader_bg": "D6DCE4",
    "meta": "FFF2CC",
}

DISCLAIMER = (
    "Bu rapor yatırım tavsiyesi değildir. Veriler TradingView canlı taramasıdır; "
    "geçmiş performans gelecekteki sonuçların göstergesi değildir."
)


# ---------------------------------------------------------------------------
# Yardımcılar
# ---------------------------------------------------------------------------

def log(msg: str) -> None:
    print(f"[Veri] {msg}", flush=True)


def norm_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip().lower()
    for a, b in (("ş", "s"), ("ğ", "g"), ("ı", "i"), ("İ", "i"),
                 ("ö", "o"), ("ü", "u"), ("ç", "c"), ("â", "a")):
        s = s.replace(a, b)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def _f(row, key):
    if key not in row.index:
        return None
    v = row[key]
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return v


def _first_valid(row, *keys):
    """Verilen anahtarlardan ilk dolu (NaN/boş olmayan) değeri döndürür."""
    for k in keys:
        v = _f(row, k)
        if v is None:
            continue
        if isinstance(v, str) and (not v.strip() or v.strip().lower() == "nan"):
            continue
        return v
    return None


# ---------------------------------------------------------------------------
# Veri çekme — canlı TradingView
# ---------------------------------------------------------------------------

def _http(url: str, data: bytes | None = None, timeout: int = 60, retries: int = 3) -> bytes:
    last = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(url, data=data, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as exc:
            last = exc
            if attempt < retries:
                time.sleep(1.2 * attempt)
    raise last


def scan_tv(columns: list[str], row_range=(0, 1000)) -> dict:
    body = {
        "symbols": {"query": {"types": []}, "tickers": []},
        "columns": list(columns),
        "range": [row_range[0], row_range[1]],
        "options": {"lang": "tr"},
    }
    payload = json.dumps(body).encode("utf-8")
    raw = _http(SCANNER_URL.format(market=MARKET), data=payload)
    return json.loads(raw.decode("utf-8", "replace"))


def fetch_live() -> pd.DataFrame:
    """TradingView scanner'dan canlı BIST verisi çeker."""
    log("canlı veri çekiliyor (TradingView scanner)...")
    probe = scan_tv(["name", "type"])
    symbols = [row["s"] for row in probe.get("data") or []]
    total = probe.get("totalCount", len(symbols))
    if not symbols:
        raise RuntimeError("Hisse listesi boş geldi.")
    log(f"{len(symbols)}/{total} sembol bulundu, {len(TV_COLUMNS)} sütun çekiliyor...")

    index = {s: i for i, s in enumerate(symbols)}
    data = {c: [None] * len(symbols) for c in TV_COLUMNS}
    batch = 80
    for start in range(0, len(TV_COLUMNS), batch):
        chunk = TV_COLUMNS[start:start + batch]
        result = scan_tv(chunk)
        for row in result.get("data") or []:
            pos = index.get(row["s"])
            if pos is None:
                continue
            values = row.get("d") or []
            for off, col in enumerate(chunk):
                if off < len(values):
                    val = values[off]
                    if isinstance(val, dict):
                        val = val.get("close") or val.get("value") or str(val)
                    data[col][pos] = val
        done = min(start + batch, len(TV_COLUMNS))
        print(f"\r       ilerleme: {done}/{len(TV_COLUMNS)}", end="", flush=True)
        time.sleep(0.08)
    print()

    rows = []
    for i, ticker in enumerate(symbols):
        row = {"symbol": ticker.split(":")[-1]}
        for internal, tv_key in TV_MAP.items():
            if tv_key is None:
                continue
            row[internal] = data.get(tv_key, [None] * len(symbols))[i]
        if isinstance(row.get("sector"), str):
            row["sector"] = SECTOR_MAP.get(row["sector"], row["sector"])
        if isinstance(row.get("industry"), str):
            row["industry"] = SECTOR_MAP.get(row["industry"], row["industry"])
        rows.append(row)

    df = pd.DataFrame(rows)
    df = clean_df(df)
    log(f"canlı veri hazır: {len(df)} hisse")
    return df


def load_offline(path: str | Path) -> pd.DataFrame:
    """Yerel Excel dosyasından okur."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Çevrimdışı veri dosyası yok: {path}")
    log(f"çevrimdışı veri okunuyor: {path.name}")
    raw = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    if len(raw) > 1 and str(raw.iloc[0, 0]).strip().lower() in ("symbol", "bist kodu"):
        if str(raw.columns[0]).strip() not in ("BIST Kodu", "symbol"):
            raw = raw.iloc[1:].reset_index(drop=True)

    header_norm = {c: norm_text(c) for c in raw.columns}
    df = pd.DataFrame()
    for tr, internal in TR_MAP.items():
        target = norm_text(tr)
        match = None
        for col, n in header_norm.items():
            if n == target or n.startswith(target[:20]):
                match = col
                break
        if match is not None:
            df[internal] = raw[match].values
    if "symbol" not in df.columns:
        df["symbol"] = raw.iloc[:, 0].values
    if "close" not in df.columns:
        for col in raw.columns:
            if norm_text(col) in ("son fiyat", "close"):
                df["close"] = raw[col].values
                break
    df = clean_df(df)
    log(f"çevrimdışı veri hazır: {len(df)} hisse")
    return df


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["symbol"] = df["symbol"].astype(str).str.strip().str.upper()
    junk = df["symbol"].isin({"SYMBOL", "BIST KODU", "KOD", "CODE", "NAME", "HISSE", "NAN"})
    if "sector" in df.columns:
        junk |= df["sector"].astype(str).str.strip().str.upper().isin({"SEKTÖR", "SEKTOR", "SECTOR"})
    df = df[~junk]
    if "close" in df.columns:
        df = df[pd.to_numeric(df["close"], errors="coerce").notna()]

    if "asset_type" in df.columns:
        at = df["asset_type"].astype(str).str.strip().str.lower()
        has_fund = pd.Series(False, index=df.index)
        for k in ("pe", "pb", "roe", "rev_growth", "market_cap"):
            if k in df.columns:
                has_fund |= pd.to_numeric(df[k], errors="coerce").notna()
        keep = (at == "stock") | ((~at.isin(["fund", "etf", "index", "type", "nan"])) & has_fund)
        keep &= ~at.isin({"fund", "etf", "index"})
        df = df[keep]

    if "company" in df.columns:
        blob = (
            df["company"].astype(str).str.lower() + " "
            + df.get("name", pd.Series("", index=df.index)).astype(str).str.lower()
        )
        bad = blob.str.contains(r"certificate|sertifika|structured|yield note|% cert", na=False, regex=True)
        df = df[~bad]

    numeric = [
        "close", "volume", "change_pct", "market_cap", "high_52w", "low_52w",
        "pe", "pb", "div_yield", "beta", "ev_ebitda",
        "roe", "roa", "roic", "gross_margin", "net_margin", "ebitda_margin",
        "current_ratio", "quick_ratio", "d_e", "net_debt_ebitda",
        "altman_z", "piotroski_f",
        "rev_growth", "ni_growth", "ebitda_growth", "fcf_growth",
        "rev_cagr5", "ni_cagr5",
        "perf_1w", "perf_1m", "perf_3m", "perf_6m",
        "perf_1y", "perf_3y", "perf_5y", "perf_ytd",
        "rec_all", "rec_ma", "rec_other", "volatility_m",
        "rel_volume", "avg_volume", "float_pct", "employees", "fcf",
    ]
    for k in numeric:
        if k in df.columns:
            df[k] = pd.to_numeric(df[k], errors="coerce")

    if "earnings_date" in df.columns:
        raw = df["earnings_date"]
        parsed = pd.Series(pd.NaT, index=raw.index, dtype="datetime64[ns]")
        numeric_d = pd.to_numeric(raw, errors="coerce")
        unix_ok = numeric_d.notna() & (numeric_d > 1_400_000_000) & (numeric_d < 2_300_000_000)
        if unix_ok.any():
            parsed.loc[unix_ok] = pd.to_datetime(numeric_d.loc[unix_ok], unit="s", errors="coerce")
        remaining = raw.notna() & parsed.isna()
        if remaining.any():
            parsed.loc[remaining] = pd.to_datetime(raw.loc[remaining], errors="coerce")
        df["earnings_date"] = parsed

    df = df.drop_duplicates(subset=["symbol"], keep="first").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Filtre
# ---------------------------------------------------------------------------

def _num_col(df: pd.DataFrame, key: str) -> pd.Series:
    """Kolon yoksa boş sayısal seri döndürür; filtreler eksik kolonda çökmesin."""
    if key in df.columns:
        return pd.to_numeric(df[key], errors="coerce")
    return pd.Series(index=df.index, dtype="float64")


def apply_filters(df: pd.DataFrame, args) -> pd.DataFrame:
    out = df.copy()
    if getattr(args, "sektor", None):
        s = args.sektor.strip().lower()
        mask = (
            out.get("sector", pd.Series("", index=out.index)).astype(str).str.lower().str.contains(s, na=False)
            | out.get("industry", pd.Series("", index=out.index)).astype(str).str.lower().str.contains(s, na=False)
        )
        out = out[mask]
    if getattr(args, "min_piyasa", None) is not None:
        out = out[_num_col(out, "market_cap").fillna(0) >= args.min_piyasa]
    if getattr(args, "max_fk", None) is not None:
        pe = _num_col(out, "pe")
        out = out[pe.isna() | (pe <= args.max_fk)]
    if getattr(args, "min_roe", None) is not None:
        out = out[_num_col(out, "roe").fillna(-999) >= args.min_roe]
    if getattr(args, "min_div", None) is not None:
        out = out[_num_col(out, "div_yield").fillna(0) >= args.min_div]
    if getattr(args, "min_growth", None) is not None:
        out = out[_num_col(out, "ni_growth").fillna(-999) >= args.min_growth]
    if getattr(args, "max_borc", None) is not None:
        out = out[_num_col(out, "net_debt_ebitda").fillna(-999) <= args.max_borc]
    if getattr(args, "min_float", None) is not None:
        out = out[_num_col(out, "float_pct").fillna(-999) >= args.min_float]
    return out


def display_name_series(df: pd.DataFrame) -> pd.Series:
    if "company" not in df.columns:
        return df["name"].astype(str)
    company = df["company"].astype(str).str.strip()
    symbol = df["symbol"].astype(str).str.strip()
    name = df["name"].astype(str) if "name" in df.columns else symbol
    use = (company.str.len() > 3) & (company.str.upper() != symbol.str.upper()) & (company.str.lower() != "nan")
    return name.where(~use, company)


# ---------------------------------------------------------------------------
# Görüntü kolonları
# ---------------------------------------------------------------------------

DISPLAY_COLS = [
    ("symbol", "Kod", 10),
    ("name", "Şirket", 36),
    ("sector", "Sektör", 22),
    ("industry", "Alt Sektör", 24),
    ("close", "Son Fiyat", 11),
    ("change_pct", "Değişim %", 10),
    ("volume", "Hacim", 14),
    ("market_cap", "Piyasa Değeri", 16),
    ("high_52w", "52H Yüksek", 12),
    ("low_52w", "52H Düşük", 12),
    ("pe", "F/K", 8),
    ("pb", "PD/DD", 8),
    ("ev_ebitda", "EV/FAVÖK", 10),
    ("div_yield", "Temettü %", 10),
    ("beta", "Beta", 7),
    ("roe", "ROE %", 9),
    ("roa", "ROA %", 9),
    ("roic", "ROIC %", 9),
    ("gross_margin", "Brüt Marj %", 11),
    ("net_margin", "Net Marj %", 11),
    ("ebitda_margin", "FAVÖK Marj %", 12),
    ("current_ratio", "Current", 9),
    ("quick_ratio", "Quick", 9),
    ("d_e", "Borç/Özkaynak", 12),
    ("net_debt_ebitda", "Net Borç/FAVÖK", 13),
    ("altman_z", "Altman Z", 9),
    ("piotroski_f", "Piotroski F", 11),
    ("rev_growth", "Gelir Büyüme %", 13),
    ("ni_growth", "Kâr Büyüme %", 12),
    ("ebitda_growth", "FAVÖK Büyüme %", 13),
    ("fcf_growth", "FCF Büyüme %", 12),
    ("rev_cagr5", "Gelir 5Y %", 11),
    ("ni_cagr5", "Kâr 5Y %", 10),
    ("perf_1w", "1H %", 8),
    ("perf_1m", "1A %", 8),
    ("perf_3m", "3A %", 8),
    ("perf_6m", "6A %", 8),
    ("perf_1y", "1Y %", 8),
    ("perf_3y", "3Y %", 8),
    ("perf_5y", "5Y %", 8),
    ("perf_ytd", "YBB %", 8),
    ("rec_all", "Teknik Tavsiye", 12),
    ("rec_ma", "Teknik MA", 11),
    ("rec_other", "Teknik Diğer", 12),
    ("volatility_m", "Volatilite %", 11),
    ("rel_volume", "Göreli Hacim", 11),
    ("avg_volume", "Ort Hacim 30G", 13),
    ("float_pct", "Halka Açık %", 11),
    ("employees", "Çalışan", 10),
    ("earnings_date", "Rapor Tarihi", 13),
    ("fcf", "Serbest Nakit Akışı", 16),
]


def to_display(df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c, _, _ in DISPLAY_COLS if c in df.columns]
    out = df[cols].copy()
    if "name" in out.columns and "company" in df.columns:
        out["name"] = display_name_series(df)
    header = {c: h for c, h, _ in DISPLAY_COLS}
    return out.rename(columns=header)


# ---------------------------------------------------------------------------
# Excel raporu
# ---------------------------------------------------------------------------

def _style_header(ws, row: int, ncols: int):
    fill = PatternFill("solid", fgColor=COLORS["header_bg"])
    font = Font(color=COLORS["header_font"], bold=True, size=10)
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border


def _write_table(ws, df: pd.DataFrame, start_row: int, widths: dict | None = None) -> int:
    if df.empty:
        ws.cell(row=start_row, column=1, value="Sonuç bulunamadı.").font = Font(italic=True)
        return start_row + 2
    # ÖNEMLİ: ws.append() her zaman son dolu satırdan SONRA yazar ve start_row'u
    # yok sayar. Bu yüzden başlık ve gövde doğrudan start_row'dan itibaren yazılır;
    # aksi halde başlık biçimi (font/dolgu/filtre/dondurma) bir satır kayar.
    rows = list(dataframe_to_rows(df, index=False, header=True))
    for i, row in enumerate(rows):
        for j, val in enumerate(row, start=1):
            ws.cell(row=start_row + i, column=j, value=val)
    header_row = start_row
    _style_header(ws, header_row, len(df.columns))
    for idx, col in enumerate(df.columns, start=1):
        if widths and col in widths:
            width = widths[col]
        else:
            sample = df[col].astype(str).head(80)
            try:
                maxlen = int(sample.map(lambda x: len(str(x))).max())
            except Exception:
                maxlen = 10
            width = max(8, min(40, maxlen + 2))
            if col in ("Şirket",):
                width = 38
            if col == "Sektör":
                width = 22
        ws.column_dimensions[get_column_letter(idx)].width = width

    last_col = get_column_letter(len(df.columns))
    last_row = header_row + len(df)
    ws.auto_filter.ref = f"A{header_row}:{last_col}{last_row}"
    ws.freeze_panes = f"B{header_row + 1}"
    return last_row + 2


def _write_meta(ws, lines: list[str], start_row: int = 1) -> int:
    r = start_row
    for i, line in enumerate(lines):
        cell = ws.cell(row=r, column=1, value=line)
        if i == 0:
            cell.font = Font(bold=True, size=13, color=COLORS["header_bg"])
        elif line.startswith("Bu rapor"):
            cell.font = Font(italic=True, size=9, color="666666")
        else:
            cell.font = Font(size=10)
            if i <= 6:
                cell.fill = PatternFill("solid", fgColor=COLORS["meta"])
        r += 1
    return r + 1


def _filter_desc(args) -> str:
    parts = []
    for attr, label in (
        ("sektor", "Sektör"), ("min_piyasa", "Min Piyasa"), ("max_fk", "Max F/K"),
        ("min_roe", "Min ROE"), ("min_div", "Min Temettü"), ("min_growth", "Min Büyüme"),
        ("max_borc", "Max Net Borç/FAVÖK"), ("min_float", "Min Halka Açık"),
        ("top", "Top N"),
    ):
        val = getattr(args, attr, None)
        if val is not None and val != "":
            parts.append(f"{label}={val}")
    return ", ".join(parts) if parts else "(varsayılan — tüm hisseler)"


def build_report(df: pd.DataFrame, args, source: str, out_dir: Path) -> Path:
    now = datetime.now()
    stamp = now.strftime("%Y%m%d")
    out_path = out_dir / f"Hisse_Verileri_{stamp}.xlsx"
    # Aynı gün tekrar çalıştırılırsa yeni dosya üretilmez; mevcut dosyanın üzerine yazılır.

    wb = openpyxl.Workbook()

    # --- Veriler (ana sayfa) ---
    ws = wb.active
    ws.title = "Veriler"
    meta = [
        f"BIST Hisse Verileri — {now.strftime('%d.%m.%Y %H:%M')}",
        f"Veri kaynağı: {source}",
        f"Hisse sayısı: {len(df)}",
        f"Filtre: {_filter_desc(args)}",
        "Not: Bu sayfa ham/normalize edilmiş verileri listeler, skor içermez.",
        DISCLAIMER,
    ]
    r = _write_meta(ws, meta)
    ws.cell(row=r, column=1, value="Hisse Listesi").font = Font(bold=True, size=12, color=COLORS["header_bg"])

    ordered = df.copy()
    if "market_cap" in ordered.columns:
        ordered = ordered.sort_values("market_cap", ascending=False, na_position="last")
    if args.top:
        ordered = ordered.head(args.top)
    disp = to_display(ordered)
    widths = {h: w for _, h, w in DISPLAY_COLS}
    _write_table(ws, disp, r + 2, widths=widths)

    # --- Sektör Özet ---
    ws2 = wb.create_sheet("Sektör Özet")
    _write_meta(ws2, [
        f"Sektör Özet — {now.strftime('%d.%m.%Y %H:%M')}",
        f"Veri kaynağı: {source}",
        "Ortalama temel oranlar (skor içermez).",
        DISCLAIMER,
    ])
    sec = sector_summary(df)
    if not sec.empty:
        _write_table(ws2, sec, 7)
    else:
        ws2.cell(row=7, column=1, value="Sektör verisi yok.")

    # --- Bilgi ---
    ws3 = wb.create_sheet("Bilgi")
    info = [
        "BIST Veri Çekici — Bilgi",
        f"Çalıştırılma tarihi/saati: {now.strftime('%d.%m.%Y %H:%M:%S')}",
        f"Veri kaynağı: {source}",
        f"Toplam hisse: {len(df)}",
        f"Filtre: {_filter_desc(args)}",
        "",
        "1) VERİ KAYNAĞI",
        "   - Canlı: TradingView scanner uç noktası (tr.tradingview.com).",
        "   - Yedek: --offline ile yerel Bist_Hisse_Tum_Veriler.xlsx okunur.",
        "",
        "2) KAPSAM",
        "   - BIST'te işlem gören hisseler (fon/ETF/sertifika ayıklanır).",
        "   - Fiyat/hacim, değerleme, kârlılık, bilanço, büyüme, performans, teknik alanlar.",
        "",
        "3) PUANLAMA",
        "   - Bu sürümde skorlama/ağırlık/not üretimi YOKTUR.",
        "   - Yalnızca çekilen veriler ve basit sektör ortalamaları listelenir.",
        "",
        "4) SAYFALAR",
        "   - 'Veriler': tüm hisselerin ham/normalize edilmiş metrikleri.",
        "   - 'Sektör Özet': sektör bazlı ortalama oranlar.",
        "   - 'Bilgi': bu sayfa.",
        "",
        "5) SÜTUNLAR",
        "   - 'Kod': BIST sembolü.",
        "   - 'Şirket': TradingView firma adı.",
        "   - Oranlar (%) olarak, mutlak değerler ham sayı olarak verilir.",
        "",
        "6) FİLTRELER",
        "   - --sektor, --min-piyasa, --max-fk, --min-roe, --min-div,",
        "     --min-growth, --max-borc, --min-float, --top",
        "",
        DISCLAIMER,
    ]
    for i, line in enumerate(info, start=1):
        ws3.cell(row=i, column=1, value=line)
    ws3.column_dimensions["A"].width = 88

    wb.active = 0
    for sheet in wb.worksheets:
        sheet.sheet_view.tabSelected = False
    wb.worksheets[0].sheet_view.tabSelected = True

    wb.save(out_path)
    return out_path


def sector_summary(df: pd.DataFrame) -> pd.DataFrame:
    if "sector" not in df.columns or df.empty:
        return pd.DataFrame()

    rows = []
    for sector_name, group in df.groupby("sector", dropna=False):
        def _mean(col, g=group):
            if col not in g.columns:
                return None
            s = pd.to_numeric(g[col], errors="coerce").dropna()
            if s.empty:
                return None
            if len(s) >= 10:
                lo, hi = s.quantile(0.01), s.quantile(0.99)
                s = s[(s >= lo) & (s <= hi)]
            return round(float(s.mean()), 2) if len(s) else None

        label = sector_name if isinstance(sector_name, str) and sector_name.strip() else "Sektörsüz"
        if isinstance(sector_name, float) and pd.isna(sector_name):
            label = "Sektörsüz"
        rows.append({
            "Sektör": label,
            "Hisse Sayısı": len(group),
            "Ort F/K": _mean("pe"),
            "Ort PD/DD": _mean("pb"),
            "Ort EV/FAVÖK": _mean("ev_ebitda"),
            "Ort ROE %": _mean("roe"),
            "Ort ROA %": _mean("roa"),
            "Ort Net Marj %": _mean("net_margin"),
            "Ort FAVÖK Marj %": _mean("ebitda_margin"),
            "Ort Gelir Büyüme %": _mean("rev_growth"),
            "Ort Kâr Büyüme %": _mean("ni_growth"),
            "Ort Net Borç/FAVÖK": _mean("net_debt_ebitda"),
            "Ort Temettü %": _mean("div_yield"),
        })

    summary = pd.DataFrame(rows)
    if summary.empty:
        return summary
    summary = summary.sort_values("Hisse Sayısı", ascending=False).reset_index(drop=True)
    return summary


def single_stock_view(df: pd.DataFrame, symbol: str) -> pd.DataFrame | None:
    """Konsol önizlemesi için tek hisse tablosu."""
    sym = symbol.strip().upper()
    row = df[df["symbol"].astype(str).str.upper() == sym]
    if row.empty:
        row = df[df["symbol"].astype(str).str.upper().str.contains(sym, na=False)]
    if row.empty:
        return None
    r = row.iloc[0]
    items = [
        ("BIST Kodu", r.get("symbol")),
        ("Şirket", _first_valid(r, "company", "name")),
        ("Sektör", r.get("sector")),
        ("Alt Sektör", r.get("industry")),
        ("Son Fiyat", r.get("close")),
        ("Günlük Değişim %", r.get("change_pct")),
        ("Piyasa Değeri", r.get("market_cap")),
        ("F/K", r.get("pe")),
        ("PD/DD", r.get("pb")),
        ("EV/FAVÖK", r.get("ev_ebitda")),
        ("ROE %", r.get("roe")),
        ("ROA %", r.get("roa")),
        ("ROIC %", r.get("roic")),
        ("Net Marj %", r.get("net_margin")),
        ("FAVÖK Marj %", r.get("ebitda_margin")),
        ("Gelir Büyüme %", r.get("rev_growth")),
        ("Kâr Büyüme %", r.get("ni_growth")),
        ("Net Borç/FAVÖK", r.get("net_debt_ebitda")),
        ("Altman Z", r.get("altman_z")),
        ("Piotroski F", r.get("piotroski_f")),
        ("Temettü %", r.get("div_yield")),
        ("1A Performans %", r.get("perf_1m")),
        ("1Y Performans %", r.get("perf_1y")),
        ("Teknik Tavsiye", r.get("rec_all")),
    ]
    return pd.DataFrame(items, columns=["Metrik", "Değer"])


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    p = argparse.ArgumentParser(
        prog="hisse_veri_cek.py",
        description="BIST verilerini TradingView'dan çekip Excel'e yazar (skorlama yok)",
        epilog=(
            "Örnekler:\n"
            "  python hisse_veri_cek.py\n"
            "  python hisse_veri_cek.py --hisse THYAO\n"
            "  python hisse_veri_cek.py --sektor Banka\n"
            "  python hisse_veri_cek.py --top 100\n"
            "  python hisse_veri_cek.py --offline\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--cikti", default=str(ARCHIVE_DIR),
                   help="Çıktı klasörü (varsayılan: script yanındaki Archive)")
    p.add_argument("--top", type=int, default=None,
                   help="İlk N hisse (piyasa değeri azalan)")
    p.add_argument("--hisse", default=None,
                   help="Tek hisse konsol detayı (örn: THYAO) — Excel'e yazılmaz")
    p.add_argument("--sektor", default=None, help="Sektör filtresi (örn: Banka)")
    p.add_argument("--min-piyasa", type=float, default=None)
    p.add_argument("--max-fk", type=float, default=None)
    p.add_argument("--min-roe", type=float, default=None)
    p.add_argument("--min-div", type=float, default=None)
    p.add_argument("--min-growth", type=float, default=None)
    p.add_argument("--max-borc", type=float, default=None)
    p.add_argument("--min-float", type=float, default=None)
    p.add_argument("--offline", action="store_true",
                   help="Canlı yerine yerel Bist_Hisse_Tum_Veriler.xlsx kullan")
    p.add_argument("--offline-dosya", default=str(DEFAULT_OFFLINE),
                   help="Çevrimdışı Excel yolu")
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    started = time.time()
    print("=" * 64)
    print("BIST Veri Çekici (skorlama yok — yalnızca veri listeleme)")
    print("=" * 64)

    # 1) Veri
    if args.offline:
        df = load_offline(args.offline_dosya)
        source = f"Çevrimdışı: {args.offline_dosya}"
    else:
        try:
            df = fetch_live()
            source = f"Canlı TradingView ({MARKET}) — {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        except Exception as exc:
            log(f"CANLI VERİ HATASI: {exc}")
            log("yerel Excel'e düşülüyor...")
            df = load_offline(args.offline_dosya)
            source = f"Çevrimdışı yedek: {args.offline_dosya} (canlı hata: {exc})"

    if df.empty:
        print("HATA: veri boş.", file=sys.stderr)
        return 1

    # 2) Filtre
    filtered = apply_filters(df, args)
    log(f"filtre sonrası: {len(filtered)}/{len(df)} hisse")

    # 3) Konsol önizleme
    preview = filtered.copy()
    if "market_cap" in preview.columns:
        preview = preview.sort_values("market_cap", ascending=False, na_position="last")
    if args.top:
        preview = preview.head(args.top)
    cols = [c for c in ("symbol", "name", "sector", "close", "market_cap",
                        "pe", "pb", "roe", "div_yield") if c in preview.columns]
    p10 = preview[cols].head(10).copy()
    if "name" in p10.columns and "company" in preview.columns:
        p10["name"] = display_name_series(preview).head(10).values
    print(f"\n  İlk {len(p10)} hisse:")
    with pd.option_context("display.max_columns", 20, "display.width", 140, "display.max_colwidth", 36):
        print(p10.to_string(index=False) if not p10.empty else "  (boş)")

    if args.hisse:
        sv = single_stock_view(filtered if not filtered.empty else df, args.hisse)
        if sv is None:
            sv = single_stock_view(df, args.hisse)
        if sv is not None:
            print(f"\n  --- {args.hisse} detay ---")
            print(sv.to_string(index=False))
        else:
            print(f"\n  UYARI: {args.hisse} bulunamadı.")

    # 4) Excel (günlük arşiv)
    out_dir = Path(args.cikti)
    out_dir.mkdir(parents=True, exist_ok=True)
    log(f"Excel raporu yazılıyor (arşiv: {out_dir})...")
    out_path = build_report(filtered if not filtered.empty else df, args, source, out_dir)
    size = out_path.stat().st_size / 1024
    log(f"Rapor hazır: {out_path.resolve()} ({size:.0f} KB)")
    log(f"süre: {time.time() - started:.1f} sn")
    print("\nBitti. Yatırım tavsiyesi değildir.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())