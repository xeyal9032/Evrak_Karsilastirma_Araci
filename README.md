# Evrak Karşılaştırma Aracı

**DATEV Buchungsstapel** ve **Journal** export dosyalarını otomatik tanır, satır satır karşılaştırır ve masaüstüne renkli Excel raporu üretir.

> TR / RU / DE / EN · GUI + CLI · HTML/PDF · batch/SQLite · 250+ otomatik test

[![CI](https://github.com/xeyal9032/Evrak_Karsilastirma_Araci/actions/workflows/ci.yml/badge.svg)](https://github.com/xeyal9032/Evrak_Karsilastirma_Araci/actions/workflows/ci.yml)

### CI nasıl çalışır?

1. **Smoke** (~2 sn): i18n/PDF sözleşmesi + match-fix + kritik paketler — erken kırmızı sinyal  
2. **Lint**: Ruff (kritik hata sınıfları)  
3. **Full matrix**: Ubuntu/Windows × Python 3.11/3.12  
4. **Publish report**: birleşik `CI_REPORT.md` artifact + Actions Step Summary; PR fail olursa otomatik yorum  

Raporu Cursor’a vermek için: Actions → ilgili run → Artifacts → **`CI_REPORT`** indirin, sohbete ekleyin.

Yerelde:

```bash
python tools/ci_runner.py --suite smoke
python tools/ci_runner.py --suite full
```

---

## Community vs Enterprise

| Özellik | Community (ücretsiz, MIT) | Enterprise (planlı, ücretli) |
|---------|---------------------------|------------------------------|
| Tek çift GUI / CLI karşılaştırma | Evet | Evet |
| Excel + HTML/PDF rapor, 4 dil, log, ilerleme | Evet | Evet |
| Klasör / toplu karşılaştırma | Evet (CLI `--batch`) | Evet + destek SLA |
| SQL sonuç arşivi | Evet (CLI `--archive`) | Evet + destek SLA |
| İmzalı ticari dağıtım / SLA | — | Planlı |

Ayrıntı: [LICENSE](LICENSE), [LICENSE.COMMERCIAL](LICENSE.COMMERCIAL), [docs/FAQ.md](docs/FAQ.md).

---

## Ne yapar?

İki muhasebe evrağı seçersiniz → sistem formatı tanır → kayıtları eşleştirir → `Karsilastirma_YYYYMMDD_HHMMSS.xlsx` raporunu açar.

| Renk | Anlam |
|------|--------|
| Sarı | İki dosyada da eşleşti |
| Turuncu | Aynı kayıt, tutar veya alan farkı — kontrol edin |
| Yeşil | Dosya içi mükerrer + storno (net sıfır, gerçek fark değil) |
| Kırmızı | Sadece bir dosyada (mevcut satır + karşı tarafta eksik) — en kritik |

Excel / HTML / PDF metinleri dil seçimine göre çevrilir (sayfa adları sabit kod).

---

## Desteklenen formatlar

| Format | Örnek |
|--------|--------|
| DATEV Buchungsstapel **CSV** | `EXTF_Buchungsstapel_*.csv` |
| DATEV Buchungsstapel **Excel** | Aynı sütunlar, `.xlsx` |
| **Journal Excel** | `Belegdat.` · `Sollkto` · `Habenkto` · `Betrag` |

Desteklenmez: PDF / tarama / OCR, önceki karşılaştırma çıktı dosyaları.

---

## Kurulum

**Gereksinim:** Python 3.9+ (Tkinter ile)

```bash
git clone https://github.com/xeyal9032/Evrak_Karsilastirma_Araci.git
cd Evrak_Karsilastirma_Araci
pip install -r requirements.txt
```

İsteğe bağlı (Windows sürükle-bırak):

```bash
pip install windnd
```

---

## Çalıştırma

### GUI

```bash
python karsilastir.py
```

### CLI (otomasyon)

```bash
python karsilastir.py tests/fixtures/sample_a.csv tests/fixtures/sample_b.csv -o out.xlsx --lang en --html --pdf --archive
python karsilastir.py tests/fixtures/sample_a.csv tests/fixtures/sample_b.csv --archive data/compare_archive.db --quiet
# Batch: iki klasorde ayni isimli dosyalari eslestirir
python karsilastir.py path\to\klasor_a path\to\klasor_b --batch -o batch_out --html --fast
python cli.py tests/fixtures/sample_a.csv tests/fixtures/sample_b.csv --quiet
```

Exit codes: `0` OK · `1` format/IO · `2` beklenmeyen hata.

**Benchmark** (ör. ~5k satır, `--fast`, bu makinede ~2.3 sn):

```bash
python benchmarks/run_benchmark.py --rows 5000 --fast
python benchmarks/run_benchmark.py --rows 100000 --fast
```

Windows’ta hazır exe (yerelde paketlerseniz):

```text
dist/Evrak_Karsilastirma_Araci.exe
dist/Evrak_Karsilastirma_Araci.exe file1.csv file2.xlsx -o out.xlsx
```

Exe windowed GUI’dir (`console=False`); CLI argümanlarıyla çalıştırınca Windows konsolu
açılır/parent’a bağlanır (stdout/stderr görünür). `logs/` ve `data/` exe yanına yazılır.

Exe paketleme:

```bash
python -m venv .venv_build
.\.venv_build\Scripts\python.exe -m pip install -r requirements.txt pyinstaller
.\.venv_build\Scripts\python.exe -m PyInstaller Evrak_Karsilastirma_Araci.spec --noconfirm
```

---

## Kullanım

1. Dil seçin (TR / RU / DE / EN)
2. **Dosya 1** ve **Dosya 2** seçin veya sürükleyip bırakın
3. Format durum çubuğunda görünür
4. **Karşılaştır** düğmesine basın — ilerleme çubuğu güncellenir
5. Rapor masaüstüne kaydedilir ve otomatik açılır

### Rapor sayfaları

- `KARSILASTIRMA-SRAVNENIE` — yan yana karşılaştırma + **Durum** sütunu (AutoFilter)
- `FILTRE-KIRMIZI` / `FILTRE-TURUNCU` — sadece kritik / değer farkı satırları
- `OZET-SVODKA` — özet ve renk açıklaması
- Dosya1 / Dosya2 — kaynak satırları
- `FARKLAR-RAZNICA` — sadece eşleşmeyenler
- `HESAP FARKI-SCHETA` — hesap bakiyesi farkları

Ekran görüntüleri için yer: [`docs/screenshots/`](docs/screenshots/).

---

## Eşleştirme mantığı (özet)

1. Dosya içi `<Storno>` + karşı kayıt netleştirmesi
2. Tam eşleşme (tutar + S/H + hesap + tarih + Belegfeld1)
3. Metin / yeniden kodlanmış hesap katmanları
4. Belegfeld1’in metne taşındığı substring eşleşmesi
5. Değer farkı (turuncu) ve şüpheli eşleşme

S/H yönü kanonikleştirilir; böylece DATEV **H** kaydı ile Journal **Soll** bakışı aynı ekonomik işlemi eşleştirebilir.

---

## Testler

```bash
python -m unittest discover -s tests -v
```

Her push / pull request’te [GitHub Actions CI](https://github.com/xeyal9032/Evrak_Karsilastirma_Araci/actions) otomatik çalışır.

Katkı: [CONTRIBUTING.md](CONTRIBUTING.md) · SSS: [docs/FAQ.md](docs/FAQ.md) · Sürümler: [CHANGELOG.md](CHANGELOG.md)

---

## Proje yapısı

```text
Evrak_Karsilastirma_Araci/
├── karsilastir.py          # GUI + CLI giriş
├── cli.py                  # CLI kısayolu
├── karsilastir_motor.py    # Algılama + karşılaştırma + Excel
├── report_extra.py         # HTML / PDF özet
├── batch_compare.py        # Klasör toplu karşılaştırma
├── archive_db.py           # SQLite arşiv
├── i18n.py / locales/      # TR RU DE EN
├── benchmarks/             # Performans ölçümü
├── docs/FAQ.md
└── tests/
```

---

## Lisans

Community Edition: [MIT](LICENSE). Enterprise özellikler için taslak: [LICENSE.COMMERCIAL](LICENSE.COMMERCIAL).

DATEV ve Journal export’larınızın yasal kullanımına siz karar verin.

---

### Document comparison tool (EN)

Compares DATEV / Journal exports, detects format automatically, and writes a colored Excel report. GUI or CLI; languages TR / RU / DE / EN.
