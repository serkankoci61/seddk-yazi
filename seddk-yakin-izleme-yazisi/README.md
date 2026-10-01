# SEDDK Yakın İzleme Yazısı Skill (v4)

DOGA Sigorta A.Ş.'nin SEDDK yakın izleme kapsamında Kuruma göndereceği **aylık raporu** ve
**olay bazlı resmi yazıları** hazırlayan agent skill'i.

## İki rejim

1. **Aylık rapor (v4'te yeni):** Kurumun E-51048673-010.05-5045010 sayılı "Kuruma Yapılacak
   Raporlamalar Hk." üst yazısı (30.09.2026) gereği, 1937 sayılı Karar md. 1 kapsamındaki
   bilgiler — 100 bin TL üzeri ödemeler, yönetim kurulu kararları, dış hizmet alımları —
   Kurumun **Veri Deseni** formatında, biten aya ilişkin, **her ayın 15'ine kadar**,
   iç denetim müdürü ve iç sistemlerden sorumlu yönetim kurulu üyesi imzalı yazı ile sunulur.
   `rapor_olustur.py` doldurulmuş xlsx + üst yazı docx birlikte üretir.
2. **Olay bazlı yazılar:** sözleşme onayı (500.000 TL üzerinde, imzadan önce), hukuk dosyası
   günlük ödeme bildirimi, 20.000.000 TL bloke hesap bildirimi, internet bankacılığı yetki
   kaldırma, düzeltme/ek bilgi. `yazi_olustur.py` ile.

## Yapı

```
seddk-yakin-izleme-yazisi/
├── SKILL.md                          # ana süreç: aylık rapor + olay bazlı yazılar
├── assets/
│   └── sirket_bilgileri.json          # muhatap, karar + raporlama künyeleri, imzacılar
├── references/
│   ├── veri_deseni.md                 # Kurum Veri Deseni sütun sütun + hücre notları (v4 yeni)
│   ├── sablonlar.md                   # yazı türleri + §11 aylık rapor üst yazısı (v4 yeni)
│   └── ornekler/                      # 10 hazır girdi JSON'u (olay bazlı türler)
└── scripts/
    ├── rapor_olustur.py               # aylık rapor: Veri Deseni xlsx + üst yazı docx (v4 yeni)
    ├── yazi_olustur.py                # olay bazlı yazı docx üretici + doğrulama
    └── dogrula.py                     # sayfa yerleşimi/yer tutucu denetleyicisi (pymupdf)
```

## Hızlı Başlangıç (aylık rapor)

```bash
# 1) Veri klasörlerini envanterle (ODEMELER/ SOZLESMELER/ YKK/)
python scripts/rapor_olustur.py --tara --veri-dizini "C:\Users\DS_COM70\Desktop\YAKIN İZLEME"

# 2) Girdi JSON'u kur (şema: rapor_olustur.py docstring; --ornek-girdi iskelet yazar),
#    ardından üret:
python scripts/rapor_olustur.py girdi.json --ustyazi
python scripts/dogrula.py SEDDK_YakinIzleme_Rapor_2026-10.docx
```

Çıktılar: `SEDDK_YakinIzleme_VeriDeseni_YYYY-AA.xlsx` + `SEDDK_YakinIzleme_Rapor_YYYY-AA.docx`.

## Hızlı Başlangıç (olay bazlı yazı)

```bash
python scripts/yazi_olustur.py references/ornekler/04_sozlesme_onay.json SEDDK_Sozlesme_Onay_Talebi.docx
python scripts/dogrula.py SEDDK_Sozlesme_Onay_Talebi.docx
```

Bağımlılıklar: `openpyxl` + `python-docx` (üretim), `pymupdf` (doğrulama).

## v4'te Yeniler (v3'e Göre)

- **Aylık rapor rejimi:** E-51048673-010.05-5045010 üst yazısı; 15'ine kadar periyodu; iç denetim
  müdürü + iç sistemlerden sorumlu YK üyesi imza şartı. Tür 1-3 tekil bildirimleri normalde
  aylık raporla yerine getirilir.
- **`rapor_olustur.py`:** Kurum Veri Deseni xlsx'ini bire bir koruyarak doldurur (başlık/tip
  satırları, hücre notları, Notlar bölümü yerinde kalır); üst yazıyı otomatik kurar. Eşik
  (100 bin TL aşan), 500 bin TL/YK-onay bağlantısı, tarih formatları, karar ayı, ek dosyası
  varlığı ve satır bütünlüğü denetimleri üremden önce/sonra çalışır.
- **`veri_deseni.md`:** 3 sayfanın sütun sütun talimatı, Kurumun hücre notları (gönderim dönemi,
  SDH Yönetmeliği md. 4 sınıflandırması, sonlanma boş bırakma kuralı, dayanak belge ek kuralı)
  ve klasör sözleşmesi (ODEMELER/SOZLESMELER/YKK).
- **Şablon §11:** aylık rapor üst yazı şablonu, boş dönem kuralı, 15 gecikme gerekçesi.

## Eşikler (1937 sayılı Karar'dan)

| Kavram | Eşik | Anlamı |
|---|---|---|
| Ödeme (rapor 01 sayfası) | 100 bin TL **üzeri** | 100.000 TL'yi aşan; eşik dahil değil; ana faaliyet kalemleri hariç |
| Sözleşme onayı (tür 4) | 500.000 TL **üzerinde** | Toplam bedel (tüm süre/taksitler); onay imzadan önce |
| Bloke tutar (tür 5/6) | 20.000.000 TL | Yalnızca hukuk dosyası ödemelerinde; günlük bildirim |
| Rapor periyodu | Aylık | Biten ay; izleyen ayın **15'i** |

## Yer Tutucu Kuralı

Bilinmeyen her değer `[●açıklama]` olarak yazılır ve docx'te **sarı vurgulanır**. Asla uydurma
veri yazılmaz; gönderimden önce yer tutucu listesi sıfırlanmış olmalıdır.