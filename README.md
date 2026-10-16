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

Workflow, **Google Apps Script'ten** günde bir kez tetiklenir (aşağıya bakın).
Çalışınca `Archive/` klasörüne yeni Excel ekler ve değişikliği otomatik olarak
repoya commit'ler. Workflow'un kendi zamanlayıcısı (cron) **yoktur**; tetikleme
yalnızca Apps Script üzerindendir.

### Kurulum adımları

```bash
git remote add origin https://github.com/fkcengizhan-crypto/BistTemelVeriler.git
git push -u origin master      # yerel branch adına göre: master veya main
```

Ardından GitHub'da:

1. **Settings → Actions → General** bölümüne gidin.
2. **Workflow permissions** kısmında **"Read and write permissions"** seçin ve kaydedin.
   (Workflow, arşivi repoya commit edebilmek için yazma iznine ihtiyaç duyar.)
3. **Actions** sekmesinden workflow'u **Run workflow** ile elle test edebilirsiniz.

### Zamanlamayı değiştirme

Zamanlama Apps Script tarafındadır. `google_apps_script.gs` içindeki
`RUN_HOUR` / `RUN_MINUTE` değerlerini değiştirip `createDailyTrigger()`
fonksiyonunu yeniden çalıştırın.

---

## Google Apps Script ile tetikleme (tek yöntem)

Tetikleme **yalnızca** Google Apps Script üzerinden yapılır. Apps Script veriyi
çekmez; sadece GitHub Actions workflow'unu (`workflow_dispatch`) başlatır.

Dosya: [`google_apps_script.gs`](google_apps_script.gs)

1. **GitHub token:** Settings → Developer settings → Personal access tokens →
   **Fine-grained tokens** → yeni token
   - Repository access: Only select repositories → **BistTemelVeriler**
   - Permissions → **Actions: Read and write**
2. `google_apps_script.gs` içindeki **`GITHUB_TOKEN`** alanına token'ı yaz:

   ```javascript
   const GITHUB_TOKEN = 'BURAYA_TOKEN_YAZ';   // <-- buraya yapıştır
   ```

3. [script.google.com](https://script.google.com) → Yeni proje →
   kodu `Code.gs` içine yapıştır
4. `triggerWorkflow()` → bir kez çalıştır (GitHub'da Actions sekmesinde yeni run görünmeli)
5. `createDailyTrigger()` → bir kez çalıştır (her gün 18:30 TRT)
6. `checkLastRun()` → son çalışmaların durumunu gösterir

> Güvenlik: Token kod içinde durduğu için Apps Script projesini başkalarıyla
> paylaşmayın. Token'ın süresi dolduğunda `GITHUB_TOKEN` alanını güncelleyin.

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
