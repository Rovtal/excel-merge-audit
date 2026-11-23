# excel-merge-audit

Dil: [English](README.md) | **Türkçe**

Birden fazla dağınık `.xlsx`, `.csv` ve `.tsv` tablosunu **%100 çevrimdışı** birleştiren, Türkçe/İngilizce sütun başlıklarını otomatik eşleyen, sayı/tarih formatlarını tekilleştiren, birebir ve çakışan mükerrer kayıtları yakalayan ve kuruş mutabakatlı Excel raporu üreten Python CLI aracı.

> Veri mahremiyeti gereği dosyalar dış ağa veya buluta gönderilmez. `openpyxl` veya `pandas` bağımlılığı gerektirmez; OpenXML `.xlsx` dosyalarını doğrudan güvenlik korumalı `zipfile + lxml` üzerinden okur, çıktıyı formüllü `XlsxWriter` dosyası olarak üretir.

## Neleri Çözer?

- **Bozuk Türkçe Karakter ve Ayırıcı Karmaşası:** `UTF-8`, `UTF-8 BOM`, `cp1254` (Türkçe Windows ANSI) ve `ISO-8859-9` dosyalarıyla `;`, `,`, `\t`, `|` ayırıcılarını otomatik çözer.
- **Üst Başlık / Boşluk / Seyrek Hücreler:** Raporun ilk `1–15` satırındaki firma başlıklarını atlar, gerçek başlık satırını bulur; atlanmış boş XML hücrelerinde sütun kaydırmaz ve Excel `1900`/`1904` tarih seri numaralarını `YYYY-MM-DD` tarihine çevirir.
- **Sıkı `Decimal` Para Ayrıştırma:** `1.250,50 TL`, `(412,50)` muhasebe eksisi ve `1,250.50 USD` biçimlerini kayıpsız kuruş hesabıyla (`Decimal`) okur; `Net + KDV = Brüt` tutarsızlıklarını `Hata` olarak işaretler.
- **Mükerrer ve Çakışan Belge Yakalama:** Birebir tekrar satırları temizleyip çıkarılan tutarı mutabakat tablosunda gösterir; aynı belge numarası ve tarafla farklı tutar/tarihte gelen satırları **Çakışan Belge (`Hata`)** olarak yakalar.
- **Güvenlik Kalkanları:** OpenXML **XXE**, **zip-slip (`../`)**, arşiv boyut sınırı ve Excel hücre formül enjeksiyonu (**CWE-1236** `=CMD|...`) korumalıdır.

## Kapsam Sınırları (Net Beyan)

- **Desteklenenler:** `.xlsx` (OpenXML), `.csv`, `.tsv` dosyaları.
- **Kapsam Dışı:** Eski ikili `.xls` (`Excel 97-2003`) ve parola korumalı/şifreli `.xlsx` dosyaları bilerek reddedilir (kullanıcıdan `.xlsx` veya `.csv` olarak kaydetmesi istenir).

## Kurulum ve Çalıştırma

Standart Python ortamında (`venv` / `pip` sağlam ise):

```bash
python -m pip install -e .
excel-merge demo --out-dir demo_output
```

Kurulum yapmadan doğrudan kaynak dizinden çalıştırmak için (PowerShell / Bash):

```powershell
$env:PYTHONPATH="src"; python -m excel_merge_audit.cli demo --out-dir demo_output
```

```bash
PYTHONPATH=src python -m excel_merge_audit.cli demo --out-dir demo_output
```

## Kendi Dosyalarınızla Kullanım

Kendi klasör ve dosyalarınızı birleştirmek için (`Hata` varsa exit code `2`; İngilizce çıktı için `--lang en`):

```bash
excel-merge merge ./dosyalar -o birlesik_rapor.xlsx --json ozet.json --fail-on-error
```

## Lisans

[MIT](LICENSE) — Rovtal
