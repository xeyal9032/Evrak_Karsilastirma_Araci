# Evrak Karşılaştırma Aracı

**DATEV Buchungsstapel** ve **Journal** export dosyalarını otomatik tanır, satır satır karşılaştırır ve masaüstüne renkli Excel raporu üretir.

> TR / RU arayüz · Windows exe veya Python ile çalışır · 170+ otomatik test

[![CI](https://github.com/xeyal9032/Evrak_Karsilastirma_Araci/actions/workflows/ci.yml/badge.svg)](https://github.com/xeyal9032/Evrak_Karsilastirma_Araci/actions/workflows/ci.yml)

---

## Ne yapar?

İki muhasebe evrağı seçersiniz → sistem formatı tanır → kayıtları eşleştirir → `Karsilastirma_YYYYMMDD_HHMMSS.xlsx` raporunu açar.

| Renk | Anlam |
|------|--------|
| Sarı | İki dosyada da eşleşti |
| Turuncu | Aynı kayıt, tutar veya alan farkı — kontrol edin |
| Yeşil | Dosya içi mükerrer + storno (net sıfır, gerçek fark değil) |
| Kırmızı | Sadece bir dosyada var — en kritik satırlar |

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

---

## Çalıştırma

```bash
python karsilastir.py
```

Windows’ta hazır exe (yerelde paketlerseniz):

```text
dist/Evrak_Karsilastirma_Araci.exe
```

Exe paketleme:

```bash
python -m venv .venv_build
.\.venv_build\Scripts\python.exe -m pip install openpyxl pyinstaller
.\.venv_build\Scripts\python.exe -m PyInstaller Evrak_Karsilastirma_Araci.spec --noconfirm
```

---

## Kullanım

1. **Dosya 1** ve **Dosya 2** seçin (`.csv` veya `.xlsx`)
2. Format durum çubuğunda görünür (DATEV CSV / Journal Excel …)
3. **Karsilastir / Сравнить** düğmesine basın
4. Rapor masaüstüne kaydedilir ve otomatik açılır

### Rapor sayfaları

- `KARSILASTIRMA-SRAVNENIE` — yan yana, kronolojik karşılaştırma
- `OZET-SVODKA` — özet ve renk açıklaması
- Dosya1 / Dosya2 — kaynak satırları
- `FARKLAR-RAZNICA` — sadece eşleşmeyenler
- `HESAP FARKI-SCHETA` — hesap bakiyesi farkları

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

Yerel masaüstü yollarına bağlı testler (ör. `excel1/mart`) CI’da yoksa atlanır (`skipUnless`); fixture tabanlı testler her ortamda koşar.

---

## Proje yapısı

```text
Evrak_Karsilastirma_Araci/
├── karsilastir.py          # Tkinter arayüz (TR/RU)
├── karsilastir_motor.py    # Algılama + karşılaştırma + Excel
├── requirements.txt
├── Evrak_Karsilastirma_Araci.spec
├── tests/                  # Birim + e2e testleri
└── .github/workflows/ci.yml
```

---

## Lisans / not

Kişisel / ofis içi muhasebe karşılaştırma aracı. DATEV ve Journal export’larınızın yasal kullanımına siz karar verin.

---

### Инструмент сравнения документов (RU)

Сравнивает два экспорта DATEV / Journal, автоматически определяет формат и создаёт цветной отчёт Excel на рабочем столе.

```bash
pip install -r requirements.txt
python karsilastir.py
```
