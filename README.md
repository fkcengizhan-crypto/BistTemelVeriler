# Günlük BIST Veri Arşivi

BIST hisselerinin temel verilerini **TradingView** canlı taramasından çeken ve
her iş günü otomatik olarak **`Archive/`** klasörüne tarihli Excel dosyası
kaydeden küçük bir proje.

- Veri çekici: [`vericekme.py`](vericekme.py)
- Otomasyon: [GitHub Actions](.github/workflows/daily-data.yml) — her iş günü 18:30 TRT
- Çıktı: `Archive/Hisse_Verileri_YYYYMMDD.xlsx` (aynı gün tekrar çalışırsa üzerine yazar)

---

## Klasör yapısı

```
.
├── vericekme.py                  # BIST verisini çeken script
├── requirements.txt              # Python bağımlılıkları
├── Archive/                      # Günlük Excel arşivi (workflow tarafından doldurulur)
│   └── Hisse_Verileri_YYYYMMDD.xlsx
├── .github/workflows/
│   └── daily-data.yml            # Günlük otomatik çalışma (GitHub Actions)
└── README.md
```

---

## Yerelde çalıştırma

```bash
# 1) Sanal ortam (önerilir)
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # macOS / Linux

# 2) Bağımlılıklar
pip install -r requirements.txt

# 3) Veriyi çek
python vericekme.py
```

Çıktı: `Archive/Hisse_Verileri_YYYYMMDD.xlsx`

### Faydalı seçenekler

```bash
python vericekme.py --hisse THYAO          # tek hisse konsol detayı
python vericekme.py --sektor Banka         # sektör filtresi
python vericekme.py --top 100              # ilk 100 hisse (piyasa değeri)
python vericekme.py --offline              # canlı yerine yerel Excel
python vericekme.py --cikti C:\Raporlar    # farklı çıktı klasörü
```

---

## Otomatik çalışma (GitHub Actions)

Workflow her **iş günü 18:30 TRT** (BIST kapanışı sonrası) çalışır,
`Archive/` klasörüne yeni Excel ekler ve değişikliği otomatik olarak repoya commit'ler.

### Kurulum adımları

```bash
git init
git add .
git commit -m "Günlük BIST veri arşivi projesi"
git branch -M main
git remote add origin https://github.com/KULLANICI/REPO.git
git push -u origin main
```

Ardından GitHub'da:

1. Depo sayfasında **Settings → Actions → General** bölümüne gidin.
2. **Workflow permissions** kısmında **"Read and write permissions"** seçin ve kaydedin.
   (Workflow, arşivi repoya commit edebilmek için yazma iznine ihtiyaç duyar.)
3. **Actions** sekmesinden workflow'u **Run workflow** ile elle test edebilirsiniz.

> Not: Zamanlanmış (`schedule`) workflow'lar yalnızca deponun **varsayılan branch**'inde
> (genelde `main`) çalışır. Workflow tetiklenmezse Actions'ın etkin olduğundan
> ve dosyanın varsayılan branch'te olduğundan emin olun.

### Zamanlamayı değiştirme

`.github/workflows/daily-data.yml` içindeki `cron` satırını düzenleyin.
Zaman **UTC** olarak verilir; TRT = UTC+3.

| İstenen saat | cron ifadesi |
|---|---|
| Her iş günü 18:30 TRT | `30 15 * * 1-5` |
| Her iş günü 21:00 TRT | `0 18 * * 1-5` |
| Her gün (hafta sonu dahil) | `30 15 * * *` |

---

## Notlar ve sınırlamalar

- Veri, TradingView'ın **resmî olmayan** tarama uç noktasından alınır. Sağlayıcı
  şartları değişebilir; IP engeli/geçici hatalar olabilir. Bu yüzden workflow
  veri çekmeyi **3 kez** dener.
- Fiyat/hacim canlı, temel oranlar TradingView'ın derlediği TTM/çeyreklik verilerdir.
- `Archive/` her gün ~340 KB büyür (yaklaşık **85 MB/yıl**). Depo boyutu sorun
  olursa eski dosyaları silebilir veya Git LFS / Actions artifact kullanabilirsiniz.
- Çıktılar **yatırım tavsiyesi değildir**.

---

## Gereksinimler

- Python 3.9+ (workflow'da 3.12 kullanılır)
- `pandas`, `numpy`, `openpyxl`
