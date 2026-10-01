# SEDDK Yakın İzleme Veri Deseni (01.09.2026 tarihli raporlama üst yazısı eki)

Kaynak: `SEDDK Yakın İzleme Formatı_01092026.xlsx` — Kurumun E-51048673-010.05-5045010 sayılı
"Kuruma Yapılacak Raporlamalar Hk." üst yazısı ekindeki Veri Deseni (3 sayfa). Bu dosya,
SEDDK'nın aylık yakın izleme raporunun **tek kabul edilebilir tablo biçimidir**; sütun adları,
sıraları ve tür bilgisi bire bir korunmalıdır. Kurum xlsx'in üstünde açıklama satırları
("Sayısal"/"Karakter"/"Tarih" tip satırı, "Notlar" bölümü) tutmuştur — bunlar Kurumun
doldurma talimatıdır, raporda yer almaz.

## Raporlama rejimi (üst yazıdan)

- **Kapsam:** 23.09.2026 tarih ve 1937 sayılı Kurul Kararı uyarınca — ana sigortacılık faaliyeti
  kapsamındaki ödemeler hariç (komisyon, hasar, reasürans, maaş, vergi ve diğer yasal
  yükümlülükler gibi) **100 bin TL üzeri ödemeler**, **yönetim kurulu kararları** ve
  **dış hizmet alımlarına** ilişkin bilgiler.
- **Periyot:** biten aya ilişkin rapor, **her ayın 15'ine kadar** düzenli olarak.
- **Biçim:** işbu yazı ekindeki şekil ve formatta (Veri Deseni), **iç denetim müdürü ve iç
  sistemlerden sorumlu yönetim kurulu üyesi tarafından imzalı yazı** ile.
- Kapsam dışı bir ayda bile rapor gönderilir; ilgili sayfada işlem yoksa satır açılmaz,
  yazıda "dönem içinde kapsam kapsamında işlem bulunmamaktadır" beyanı verilir (bkz. sablonlar.md §11).

## Genel kurallar (3 sayfa için ortak)

- **Yıl / Ay** (sayısal): bilginin **gönderim dönemi** (hücre notu: "Denetim 2: Bilginin
  gönderim dönemi."). Yıl = 2026 gibi 4 haneli; Ay = 1-12. Ödeme tablosunda ayrıca
  "Ödeme Ayı" sütununda ödemenin ait olduğu ay yazılır; gönderim dönemi ayı ile
  farklıysa karışıklığı önlemek için K sütunu ödeme ayını taşır.
- **Şirket Adı** (karakter): "DOGA SİGORTA ANONİM ŞİRKETİ" (tam unvan).
- **Şirket Kodu** (sayısal): SEDDK şirket kodu — `sirket_bilgileri.json → sirket_kodu`'dan.
  Bilinmiyorsa kullanıcıya SOR (yer tutucu bırak, uydurma).
- **Para birimi:** tutar sütunları sayısal; raporun yazı kısmında TL varsayılır, TL değilse
  ayrıca yazıda belirtilir.
- Hücre notları Kurumun resmî beklentisini fiksler — doldururken notla çelişen değer yazma.

## Sayfa 1: `01-100+` — 100 bin TL üzeri ödemeler

Satır 1 başlıklar (satır 2 tip bilgisi; raporda tip satırı YOK):

| Sütun | Başlık | Tür | Hücre notu / doldurma kuralı |
|---|---|---|---|
| A | Yıl | Sayısal | Gönderim dönemi yılı |
| B | Ay | Sayısal | Gönderim dönemi ayı (1-12) |
| C | Şirket Adı | Karakter | Tam unvan |
| D | Şirket Kodu | Sayısal | SEDDK şirket kodu |
| E | Ödemenin Yapıldığı Kişi/Kurum vb. | (serbest) | Alıcı unvanı; gerçek kişiyse KVKK gereği yalnızca ad-soyad (TC kimlik/IBAN yazma) |
| F | İşlem Açıklaması | Karakter | Not: "İşlem içeriğinin özet olarak verilmesi" — fatura/sözleşme no vb. ile birlikte kısa özet |
| G | Dayanak Belge | (serbest) | Fatura/sözleşme/YK kararı/dekont referansı |
| H | Hesap Kodu | Sayısal | Not: "Muhasebeleştirilen hesap kodu" (MSUG hesap planı kodu) |
| I | Hesap Adı | Karakter | MSUG hesap planı adı |
| J | Ödeme/Para Çıkış Tutarı | Sayısal | TL; 100.000,00 üzeri (eşik dahil değil) |
| K | Ödeme/Çıkış tarihi | (tarih) | GG.AA.YYYY önerilir |
| L | Ödeme Ayı | (sayısal) | Ödemenin gerçekleştiği ay (gönderim dönemiyle farklıysa ayrık gösterir) |
| M | Dayanak Belge (Ek) | — | Not: "HOE Denetim: Dayanak belge ek olarak iletilecek." → ek listesinde |

Kapsam notları:
- Eşik: J > 100.000,00 (100 bin TL **üzeri**; tam 100.000,00 kapsam dışı).
- Ana faaliyet içi kalemler (komisyon, hasar, reasürans, maaş, vergi, yasal yükümlülükler) rapora ALINMAZ.
- Hukuk dosyası ödemeleri tür 5'in günlük bildirim rejimine tabidir; bu tabloya ayrı satır olarak
  ALINMAZ, yazıda atıf yapılır (karar md. 4). (Bir günlük bildirim yazısı ayrıca gönderilir.)
- Internet bankacılığı transferi (kendi hesaplar arası hariç) karar gereği yoktur; ödeme kanalı
  tabloda yer almasa da dayanak belge ekinde görülür — tutarlılığı kontrol et.

## Sayfa 2: `02-Dış Hizmet` — Dışarıdan hizmet alımlarına ilişkin sözleşmeler

Sayfa başlığı hücresi (A1): "2. Dışarıdan hizmet alımlarına ilişkin sözleşmeler."
Satır 2 başlıklar (satır 3 tip bilgisi; A6 "Notlar", A7 not metni — bunlar raporda YOK):

| Sütun | Başlık | Tür | Hücre notu / doldurma kuralı |
|---|---|---|---|
| A | Yıl | Sayısal | Gönderim dönemi yılı |
| B | Ay | Sayısal | Gönderim dönemi ayı |
| C | Şirket Adı | Karakter | Tam unvan |
| D | Şirket Kodu | Sayısal | SEDDK şirket kodu |
| E | Hizmet Konusu | Karakter | Not: **Sigortacılık Destek Hizmet Hakkında Yönetmelik md. 4**'e göre sınıflandırma; uymayan konular da ayrıca belirtilir — bkz. aşağıdaki md.4 sınıflandırma listesi |
| F | Alınan Hizmetin Özet İçeriği | (serbest) | Kapsamın kısa betimi |
| G | Hizmetin Alınma Amacı | Karakter | Gerekçe |
| H | Hizmet Bedelinin Hesaplanma Şekli | "Üretime Bağlı/Hasara Bağlı/…'a bağlı/Belirli Bedel" | Not: sözleşmede belli ise o tutar; faydalanmaya bağlı ise tahmini tutar belirtilir |
| I | Birim Maliyeti | (serbest) | Not: "Sözleşmenin niteliğine göre varsa birim maliyetini yazınız" — yoksa boş/"—" |
| J | Hizmet Sağlayıcısının Adı | Karakter | Unvan |
| K | Ek | "Sözleşmenin PDF hali" | Sözleşme PDF'i ek olarak iletilir |
| L | Sözleşme Başlangıç Tarihi | Tarih | Not: "Sözleşmede belirlenen başlama tarihi." |
| M | Sözleşme Bitiş Tarihi | Tarih | Not: "Sözleşmede belirlenen bitiş tarihi." |
| N | Sözleşme Sonlanma Tarihi | Tarih | Not: "Sözleşme devam ediyorsa boş bırakılacak." |
| O | Sözleşmenin Toplam Bedeli | (sayısal) | Not: "Sözleşmede belirtilen bedel veya sözleşme türüne göre tahmini bedel." |
| P | Tahakkuk Eden Hizmet Bedeli | (sayısal) | Dönem içinde tahakkuk eden |
| Q | Ödenen Kısım | (sayısal) | Dönem içinde ödenen |
| R | Ödemenin Kaydedildiği Hesap Kodu | (sayısal) | MSUG hesap planı kodu |
| S | YK Kararı Alındı mı? | Karakter | Evet/Hayır (+ tarih-sayı varsa) |
| T | Ek | "YK Kararının PDF hali" | YK kararı PDF'i ek olarak iletilir |

### E sütunu — md.4 sınıflandırma listesi (kurallar)

SEDDK hücre notu (birebir): *"Sigortacılık Destek Hizmet Hakkında Yönetmeliğin 4'üncü maddesine
göre sınıflandırılmalıdır. Söz konusu sınıflandırmaya uymayan konular da ayrıca olarak
belirtilmelidir."* Rapor yazımında E sütununa `MD4(<bent>) — <bent adı>` etiketi yazılır
(örn. `MD4(ğ) — Çağrı merkezi hizmetleri`); script etiketin geçerliliğini denetler.

| Kod | Sınıf (Yönetmelik md.4 bent metni) |
|---|---|
| MD4(a) | Poliçe tanzimi ile tazminat tedvir ve ödenmesine ilişkin süreçlerde, sigorta eksperliği işinden ayrı olmak kaydıyla gerçekleştirilen teknik inceleme ve kontrol hizmetleri |
| MD4(b) | Hasar öncesi risk azaltmaya ve hasar sonrası zarar azaltmaya yönelik hizmetler |
| MD4(c) | Hasar ihbarı alma, dosya açma ve tamamlama hizmetleri |
| MD4(ç) | Onarım ve bakım hizmetleri |
| MD4(d) | Yedek parça tedarik ve kontrol hizmetleri |
| MD4(e) | Yardım (asistans) hizmetleri |
| MD4(g) | Tedavi ve yardım hizmetleri |
| MD4(ğ) | Çağrı merkezi hizmetleri |
| MD4(h) | Sovtaj yönetimi hizmetleri |
| MD4(ı) | Rücu takip hizmetleri |
| MD4(i) | Arşiv yönetimi hizmetleri |
| MD4(j) | Ürün ve tarife hazırlama hizmetleri |
| MD4(BS) | Bilgi sistemleri (md.4/2: yönetim, içerik tasarımı, erişim, kontrol, denetim, güncelleme, bilgi/rapor alma fonksiyonlarında karar gücü şirkette kalmak şartıyla) |
| KD1 | Personel istihdamına ilişkin hizmet alımları (Yön. md.1/2-a) — kapsam dışı, raporda AYRICA |
| KD2 | Avukatlık + vergi/hukuk danışmanlığı dâhil her türlü danışmanlık (Yön. md.1/2-b) — kapsam dışı, AYRICA |
| KD3 | Reklam faaliyetleri (Yön. md.1/2-c) — kapsam dışı, AYRICA |

Dikkat:
- **(f) bendi yok**: "Sigortacılık hasar tedvir uygulamalarında tıbbi danışmanlık hizmetleri"
  Danıştay 10. Daire 17/5/2021, E.:2016/1892; K.:2021/2273 kararıyla iptal — bu hizmeti
  MD4(g) altında değerlendir.
- Birden fazla bent kapsamına giren hizmette en baskın sınıf seçilir; diğerleri F sütununda
  belirtilir.
- KD sınıfları Yönetmeliğin md.1/2'sinde kapsam dışı tutulan hizmetlerdir; SEDDK hücre notu
  gereği raporda ayrıca belirtildiklerinden formlarda toplanır ve E sütununa KD etiketiyle
  yazılır (amacı G sütununda açıkça gerekçelendirilir).

Sayfa 2 kapsam notu (A7'den): **"2026 yıl mali tablolarında giderleştirilmiş tüm sözleşmeler.
(Ödenen veya daha önce ödenmekle birlikte, 2026 yılı mali tablolarına yansıtılan giderlere konu
sözleşmeler)."** — yani bu sayfa yalnız yeni imzalananları değil, 2026 mali tablolarına gider
yansıtılan TÜM dış hizmet sözleşmelerini her raporda (kümülatif güncel duruma uygun) içerir.

500.000 TL kontrolü: O sütunu (toplam bedel) 500.000 TL'yi aşıyorsa ve ana faaliyet dışıysa,
Kurum onayı (tür 4 yazısı) gereklidir; S sütunu bu onayın dayanağını gösterir.

## Sayfa 3: `03- YKK` — Yönetim Kurulu Kararları

Sayfa başlığı hücresi (A1): "İlgili ay içerisinde alınan Yönetim Kurulu kararları"
Satır 3 başlıklar (satır 4 tip bilgisi; raporda YOK):

| Sütun | Başlık | Tür | Doldurma kuralı |
|---|---|---|---|
| A | Yıl | Sayısal | Gönderim dönemi yılı |
| B | Ay | Sayısal | Gönderim dönemi ayı |
| C | Şirket Adı | Karakter | Tam unvan |
| D | Şirket Kodu | Sayısal | SEDDK şirket kodu |
| E | Karar Tarihi | Tarih | GG.AA.YYYY |
| F | Karar Sayısı | Sayısal | YK karar numarası |
| G | Gündem Konu Başlıkları | Karakter | Gündem maddesinin konu başlığı |
| H | Kabul Edilen Kararlar | Karakter | Kararın özeti/sonucu |
| I | Ek | "YK Kararının PDF hali" | YK kararı PDF'i ek olarak iletilir |

Karar "yönetim kurulu kararları" derken ayrım yapmadığından (tipik yorum) **ilgili ayda alınan
tüm YK kararları** listelenir; kullanıcı aksini teyit etmedikçe ayrım yapılmaz.

## Rapor paketi yapısı

Her ayın raporu üç parçadan oluşur:
1. **Üst yazı (.docx):** aylık rapor sunum yazısı (bkz. `sablonlar.md` §11) — Kurum üst yazısına
   İlgi, dönem beyanı, imzacılar (iç denetim müdürü + iç sistemlerden sorumlu YK üyesi).
2. **Veri deseni (.xlsx):** Kurum formatının doldurulmuş hali — 3 sayfa, sütun yapısı bire bir.
3. **EKLER:** dayanak belgeler (ödeme dekontu/fatura), sözleşme PDF'leri (K2), YK kararı PDF'leri
   (T2, I4) — sıra, Excel'deki satır numarasına göre Ek-1, Ek-2... diye numaralanır.

Dosya adları (teslimde):
- `SEDDK_YakinIzleme_Rapor_YYYY-AI.docx` (üst yazı; AI = ayın iki haneli rakamı, örn. 2026-10)
- `SEDDK_YakinIzleme_VeriDeseni_YYYY-AI.xlsx`
- `SEDDK_YakinIzleme_Ekler_YYYY-AI/` (PDF ekler; 01_odeme_…, 02_sozlesme_…, 03_yk_… alt sırasıyla)

## Klasör sözleşmesi (raporu kurmak için girdi verisi nereden okunur)

Kullanıcının veri klasörleri `C:\Users\DS_COM70\Desktop\YAKIN İZLEME` altında:

```
YAKIN İZLEME/
├── ODEMELER/      # ay içinde 100 bin TL aşan ödemelerin dayanak belgeleri (fatura, dekont, YK kararı PDF vb.)
├── SOZLESMELER/   # dış hizmet sözleşmeleri + varsa YK kararı/onay yazısı PDF'leri
├── YKK/           # yönetim kurulu kararları PDF'leri
└── SEDDK Yakın İzleme Formatı_01092026.xlsx  # boş Kurum formatı (şablon)
```

- Her belge PDF olarak konur; dosya adında tarih ve konu bulunmalıdır (örn. `2026-10-05_TeknikHizmet_Sozlesmesi.pdf`).
- Aynı işlem için birden çok belge varsa klasörde alt klasör açılabilir.
- Rapor üretimi `scripts/rapor_olustur.py` ile: klasörleri tarar, kullanıcıdan gelen ek
  metaveri (tutarlar, hesap kodları vb.) ile Excel'i doldurur ve üst yazıyı üretir (aşağıda İş Akışı).