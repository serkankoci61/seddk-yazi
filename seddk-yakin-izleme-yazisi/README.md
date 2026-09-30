# SEDDK Yakın İzleme Yazısı Skill (v3)

DOGA Sigorta A.Ş.'nin SEDDK'nın 23.09.2026 tarih ve 1937 sayılı Kurul Kararı (yakın izleme)
kapsamında Kuruma göndereceği resmi yazıları hazırlayan agent skill'i.

## Yapı

```
seddk-yakin-izleme-yazisi/
├── SKILL.md                          # ana süreç, kapsam kontrolleri, doğrulama disiplini
├── assets/
│   └── sirket_bilgileri.json          # muhatap, karar künyesi (E-97354901-045.01-5033535 / 1937), imzacılar
├── references/
│   ├── sablonlar.md                   # 9 yazı türü + birleşik yazı şablonları, eşik tablosu
│   └── ornekler/                      # 10 hazır girdi JSON'u (her tür + birleşik)
│       ├── 01_odeme_bildirimi.json
│       ├── 02_yk_karari.json
│       ├── 03_dis_hizmet.json
│       ├── 04_sozlesme_onay.json
│       ├── 05_hukuk_gunluk.json
│       ├── 06_hesap_bildirimi.json
│       ├── 07_internet_bankaciligi.json
│       ├── 08_duzeltme.json
│       ├── 09_periyodik_ozet.json
│       └── 10_birlesik.json
└── scripts/
    ├── yazi_olustur.py                # docx üretici + kapsamlı girdi doğrulama (python-docx)
    └── dogrula.py                     # sayfa yerleşimi/yer tutucu denetleyicisi (pymupdf)
```

## Hızlı Başlangıç

```bash
python scripts/yazi_olustur.py references/ornekler/01_odeme_bildirimi.json SEDDK_Odeme_Bildirimi_30092026.docx
python scripts/dogrula.py SEDDK_Odeme_Bildirimi_30092026.docx
```

Bağımlılıklar: `python-docx` (üretim), `pymupdf` (doğrulama; yoksa XML moduna düşer).

## Kullanım

1. SKILL.md içindeki tür tablosundan yazı türünü seç (1 ödeme, 2 YK kararı, 3 dış hizmet,
   4 sözleşme onayı, 5 hukuk günlük, 6 hesap bildirimi, 7 internet bankacılığı, 8 düzeltme,
   9 periyodik özet; çok işlem = 10 birleşik kuralı).
2. `references/ornekler/` altındaki ilgili örneği kopyala, gerçek verilerle uyarla.
3. `yazi_olustur.py` ile üret — script eşik biçimlerini, tablo aritmetiğini, İlgi/Ek atıf
   uyumunu doğrular; uyuşmazlıkta üretimi durdurur.
4. `dogrula.py` ile sayfa yerleşimini denetle (3+ sayfa, tek kalan imza bloğu, yer tutucular).
5. Yer tutucuları (`[●...]`, sarı vurgulu) doldur, e-imza/KEP sürecine ver.

## v3'te Yeniler (v2'ye Göre)

- **Karar metni tek doğruluk kaynağı:** Eşikler ve kapsamlar kararın birebir metnine sabitlendi
  (SKILL.md'de tablo halinde); `sirket_bilgileri.json`'a tam karar künyesi eklendi.
- **Üretim öncesi zorunlu doğrulamalar:** tablo toplam satırı **aritmetik kontrol**,
  İlgi (x) atıf ↔ ilgi listesi uyumu, Ek-N atıf ↔ ek listesi uyumu, Türk para formatı,
  kapanış cümlesi, karar tarihinden önceki yazı tarihi uyarısı.
- **Tek ilgi kuralı:** resmî yazışmaya uygun olarak tek ilgi harflenmez; atıf "İlgi yazınız ile".
- **Sayfa düzeni:** 1,15 satır aralığı (gerçek örnek yazıyla bire bir), kapanış→imza→Ek bloğu
  bölünmeye karşı bağlı, imza bloğu tek parça; `dogrula.py` ile otomatik yerleşim denetimi.
- **Tam örnek kapsamı:** 9 türün tamamı + birleşik yazı örneği (v2'de 5/9).
- **Dosya adlandırma standardı:** `SEDDK_<TurAdi>_<GGAAYYYY>.docx`.
- **Ortam bağımsız:** Claude "docx skill" atıfları kaldırıldı; doğrulama kendi scriptine taşındı.

## Yazı Türleri ve Eşikler

| Tür | Eşik (Karar'dan) |
|---|---|
| 1 Ödeme bildirimi | 100 bin TL **üzeri**, ana faaliyet dışı; komisyon/hasar/reasürans/maaş/vergi/yasal hariç |
| 4 Sözleşme onayı | 500.000 TL **üzerinde** toplam bedel; onay imzadan önce |
| 5/6 Hukuk | Günlük bildirim; 20.000.000 TL bloke yalnızca hukuk dosyasında |
| 7 İnternet bankacılığı | Transfer yetkileri tümüyle kaldırılır (bildirim isteğe bağlı) |

## Yer Tutucu Kuralı

Bilinmeyen her değer `[●açıklama]` olarak yazılır ve docx'te **sarı vurgulanır**.
Asla uydurma veri yazılmaz; `yazi_olustur.py` çıktısındaki yer tutucu listesi
gönderimden önce sıfırlanmış olmalıdır.