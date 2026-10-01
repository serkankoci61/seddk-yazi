---
name: seddk-yakin-izleme-yazisi
description: Use when the user needs any official letter or periodic report to SEDDK under the 23.09.2026 date and 1937-numbered Board decision putting DOGA Sigorta A.Ş. under close surveillance (yakın izleme). Covers the MONTHLY close-surveillance report (Veri Deseni xlsx + cover letter, due by the 15th of each month, signed by the internal audit manager and the board member responsible for internal systems, per E-51048673-010.05-5045010), plus event-driven letters — contract approval requests above 500.000 TL, daily law-file payment notifications, the 20.000.000 TL blocked-amount account notification, internet banking transfer-authority removal, correction/additional-info letters. Also triggers on "SEDDK'ya yazı", "aylık rapor", "veri deseni", "Kuruma bildirim", "yakın izleme raporu", "hukuk ödemesi bildirimi", "düzeltme yazısı". Produces a Konu/İlgi/Açıklama/Ek structured, dual-signed Word (.docx) letter and the filled Kurum-format Excel (.xlsx), with pre-flight scope and arithmetic validation.
version: 4.0.0
author: DOGA Sigorta İç Denetim / Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [seddk, yakin-izleme, resmi-yazi, doga-sigorta, bildirim, kurum-onayi, aylik-rapor, veri-deseni]
    related_skills: [audit-report-generation]
---

# SEDDK Yakın İzleme Yazısı ve Aylık Rapor Hazırlama (v4)

## Genel Bakış

SEDDK, Şirketi E-97354901-045.01-5033535 sayılı üst yazısıyla 23.09.2026 tarih ve 1937 sayılı
Kurul Kararı uyarınca yakın izleme kapsamına almıştır. Ardından Kurum, **E-51048673-010.05-5045010**
sayılı "Kuruma Yapılacak Raporlamalar Hk." üst yazısıyla (30.09.2026, Dilek SAKALLIOĞLU – Başkan a.)
raporlama biçimini belirlemiştir: Karar md. 1 kapsamındaki bilgiler (100 bin TL üzeri ödemeler,
YK kararları, dış hizmet alımları) **Kurumun Veri Deseni formatında, biten aya ilişkin, HER AYIN
15'İNE KADAR**, Şirketin **iç denetim müdürü ve iç sistemlerden sorumlu yönetim kurulu üyesi**
tarafından imzalı yazı ile sunulur.

Bu skill iki rejimi tek tutarlı süreçle yönetir:

**A. Aylık rapor (ana yükümlülük):** tür belirle → veri klasörlerini envanterle → eksikleri tek
seferde sor → girdi JSON'u kur → `rapor_olustur.py` → `dogrula.py` → teslim et.
**B. Olay bazlı yazılar:** sözleşme onayı (4), hukuk günlük bildirimi (5), hesap bildirimi (6),
internet bankacılığı (7), düzeltme (8). Bunlar aylık rapora girmez; kendi zamanında ayrı yazıyla
gider.

Resmî yazışmada tek bir kusur (yanlış İlgi, eksik ek, aritmetik hata, taşan tablo, format dışı
Excel) Kurumun geri dönüşüne yol açar. Bu yüzden bu skill yalnızca şablon değil, **yazıdan önce ve
sonra zorunlu kontrol disiplini** tanımlar. Veri uydurma yok; eksikler `[●açıklama]` yer tutucusu
olarak işaretlenir ve script bunları sarı vurgular.

## Kararın birebir metni (tek doğruluk kaynağı)

Tüm eşik ve kapsam kuralları bu metne dayanır; başka yerden eşik okuma, kararın koyduğu hükümleri
yorumla genişletme. Karar üst yazısı (E-97354901-045.01-5033535) dört karardan oluşur:

| # | Karar (birebir) | Nasıl yerine getirilir |
|---|---|---|
| 1 | "Şirketinizin yakın izleme kapsamına alınmasına **ve** ana sigortacılık faaliyeti kapsamındaki ödemeler hariç (komisyon ödemeleri, hasar ödemeleri, reasürans ödemeleri, maaş ödemeleri, vergi ve diğer yasal yükümlülükler gibi) **100 bin TL üzeri ödemeler**, **yönetim kurulu kararları** ve **dış hizmet alımlarına** ilişkin bilgileri yakın izleme kapsamında Kurumumuza göndermesine" | **AYLIK RAPOR** (Veri Deseni 3 sayfası: `01-100+`, `02-Dış Hizmet`, `03- YKK`) |
| 2 | "Sigortacılık ana faaliyeti dışında yükümlülük doğuran ve tutarı **500.000 TL üzerinde** kalan tüm sözleşmelerin Kurumumuz **onayına** tabi tutulmasına" | Olay bazlı: **sözleşme onay talebi yazısı** (imzadan ÖNCE) |
| 3 | "Şirketinizin internet bankacılığı aracılığıyla para transferi (Şirketinizin **kendi hesapları arasındaki işlemler hariç**) yetkilerinin meblağ, banka, finansal kuruluş ve konudan bağımsız olarak **tümüyle kaldırılmasına**" | Olay bazlı: uygulama bildirim yazısı (belgeleme amaçlı) |
| 4 | "Hukuk dosyaları kapsamında yapılan ödemelerin **günlük** olarak Kurumumuza bildirilmesi ve serbest bırakılacak tutarın **sadece hukuk dosyalarında kullanılması** kaydıyla, Şirketiniz tarafından belirlenecek hesapta tutulacak **20.000.000 TL'lik bloke tutarının serbest bırakılmasına** | Olay bazlı: **hesap bildirimi** (ilk ödemeden önce) + **günlük ödeme bildirimi** |

Künyeler: `assets/sirket_bilgileri.json` → `karar_kunyesi` (1937 Karar üst yazısı) ve
`raporlama_kunyesi` (E-51048673-010.05-5045010 raporlama üst yazısı; imzacıları, periyodu, format
dosyası adı). Gövdede atıf her zaman bu künyelerde bire bir yazılan üst yazı ve Karar ile yapılır.

## Aylık rapor rejimi (E-51048673-010.05-5045010)

- **Kapsam:** md. 1 bilgileri — ana sigortacılık faaliyeti kapsamındaki ödemeler hariç 100 bin TL
  **üzeri** ödemeler, yönetim kurulu kararları, dış hizmet alımlarına ilişkin sözleşmeler.
- **Periyot:** biten aya ilişkin; **her ayın 15'ine kadar** (örn. Ekim raporu → 15 Kasım).
- **Biçim:** Kurum Veri Deseni (`SEDDK Yakın İzleme Formatı_01092026.xlsx`) — 3 sayfa, sütun
  yapısı **bire bir**; sütun adları/sıraları değiştirilmez, Kurumun tip satırı ve Notlar bölümü
  raporda yer almaz. Sütun sütun talimat ve hücre notları: `references/veri_deseni.md` (okumadan
  doldurma).
- **İmzacılar:** iç denetim müdürü **ve** iç sistemlerden sorumlu yönetim kurulu üyesi (raporlama
  üst yazısının şartı; genel yazışmadaki Müdür/GMY ikilisinden FARKLI). Serkan KOÇ'un iç denetim
  müdürü olduğu varsayılır; iç sistemlerden sorumlu YK üyesinin adı kullanıcıdan teyit edilir.
- **Ekler:** dayanak belgeler (ödeme/fatura), sözleşme PDF'leri, YK kararı PDF'leri — Veri
  Deseni'nin M/K/T/I sütunlarında atıf yapılan her belge ek olarak iletilir.
- İşlem olmayan ayda bile rapor gönderilir (boş sayfa + üst yazıda "işlem yoktur" beyanı).

### Veri klasörleri (kullanıcının veri girişi)

`C:\Users\DS_COM70\Desktop\YAKIN İZLEME\` altında:

```
ODEMELER/      → 100 bin TL aşan ödemelerin dayanak belgeleri (fatura, dekont, sözleşme, YK kararı PDF)
SOZLESMELER/   → dış hizmet sözleşmeleri + varsa YK kararı / Kurum onay yazısı PDF'leri
YKK/           → yönetim kurulu kararları PDF'leri
```

Kullanıcı belgeleri bu klasörlere koyar; rapor bu klasörlerdeki verilere göre hazırlanır.
Dosya adlarında tarih ve konu bulunmalı (örn. `2026-10-05_TeknikHizmet_Sozlesmesi.pdf`).
Klasördeki belgelerden okunamayan metaveri (tutar, hesap kodu, tarih) eksiklerini **tek seferde
kullanıcıya sor**; belge adlarından çıkarılabiliyorsa taslak değerle işaretle.

### Aylık rapor iş akışı

1. **Envanter:** `python scripts/rapor_olustur.py --tara --veri-dizini "<YAKIN İZLEME yolu>"` —
   üç klasörün içeriğini listeler; yeni belgeleri kullanıcıyla gözden geçir.
2. **Türü belirle:** her belgeyi Veri Deseni sayfalarından birine eşle (ödeme→01, sözleşme→02,
   YK kararı→03). Hukuk dosyası ödemesi 01'e ALINMAZ (günlük rejim, tür 5). Ana faaliyet içi
   kalem (maaş, vergi, komisyon, hasar, reasürans) rapora ALINMAZ.
3. **Eksikleri tek seferde sor:** tutarlar, hesap kodları, tarihler, şirket kodu, iç sistemlerden
   sorumlu YK üyesi adı. Kullanıcının verdiği bilgiyi bir daha sorma.
4. **Girdi JSON'u kur** (şema `scripts/rapor_olustur.py` docstring'inde): `odemeler`,
   `sozlesmeler`, `ykk` listeleri; her satırın ek dosyası klasörde aranır.
5. **Üret ve doğrula:**
   ```bash
   python scripts/rapor_olustur.py girdi.json --ustyazi
   python scripts/dogrula.py SEDDK_YakinIzleme_Rapor_YYYY-AA.docx
   ```
   Script üretime geçmeden şunları denetler: J sütunu eşiği (100 bin TL aşan), 500 bin TL aşan
   sözleşmede YK kararı/onay kontrolü, tarih formatları, karar ayının gönderim dönemine aitliği,
   ek dosyaların klasörde varlığı, sonlanma tarihi kuralı (sözleşme sürüyorsa boş), üst yazı
   tarihinin 15'ine sığması. Üretim sonrası xlsx'i yeniden açıp satır sayılarını ve eşikleri
   doğrular.
6. **Teslim:** dosya adları `SEDDK_YakinIzleme_VeriDeseni_YYYY-AA.xlsx` +
   `SEDDK_YakinIzleme_Rapor_YYYY-AA.docx`; teslim notunda dönem, satır sayıları, toplamlar,
   eksik yer tutucular ve kapsam uyarıları yer alır.

## Olay bazlı yazılar

Tür 1-3 (tekil ödeme/YK/dış hizmet bildirimi) **normalde aylık raporla yerine getirilir**;
yalnızca Kurum özel olarak istediyse veya kullanıcı ayrı yazı isterse `references/sablonlar.md`'den
hazırlanır. Sürekli gerekli olanlar:

- **Sözleşme onayı (tür 4):** 500.000 TL üzerinde toplam bedelli ana faaliyet dışı sözleşme —
  imzadan ÖNCE. Aylık raporun 02 sayfasında bu sözleşme göründüğünde onayın alınıp alınmadığını
  denetle.
- **Hukuk günlük bildirimi (tür 5) + hesap bildirimi (tür 6):** günlük rejim; bloke 20.000.000 TL.
- **İnternet bankacılığı (tür 7), düzeltme (tür 8):** belgeleme ve tashih.

Şablonlar: `references/sablonlar.md` (§4-§8); hazır girdi örnekleri `references/ornekler/`.

## Zorunlu yazı yapısı (ev formatı)

Biçim, Şirketin Kuruma gönderdiği gerçek yazılardan alınmıştır; `yazi_olustur.py` üretir:

1. **Muhatap (solda, kalın, 3 satır):** SİGORTACILIK VE ÖZEL EMEKLİLİK / DÜZENLEME VE DENETLEME
   KURUMU / İSTANBUL. Sağda kalın **Tarih: GG.AA.YYYY** ve **Ref: YYYY/NNN**. Antetli kâğıt için
   üstte boşluk bırakılır.
2. **Konu:** tek cümlelik başlık (aylık raporda: `Yakın İzleme Kapsamında <Ay> <Yıl> Dönemine
   İlişkin Rapor Sunumu`).
3. **İlgi:** (a) her zaman 1937 Karar üst yazısı; aylık raporda (b) E-51048673-010.05-5045010
   raporlama üst yazısı. Birden çok ilgi a), b) harflenir; **tek ilgi harflenmez**; gövdedeki
   atıf "İlgi yazınız ile" / "İlgi (b) yazınız ile" uyumu scriptçe denetlenir.
4. **Açıklama:** iki yana yaslı, ilk satır girintili, 1,15 aralıklı. İlk paragrafta amaç; en sonda
   girintisiz kapanış ("Bilgilerinize arz ederiz." vb. — script denetler).
5. **Saygılarımızla, / DOGA SİGORTA A.Ş.** + yan yana iki imza. Aylık raporda imzacılar
   raporlama künyesinden: iç denetim müdürü + iç sistemlerden sorumlu YK üyesi. Diğer yazılarda
   varsayılan Müdür/GMY ikilisi (JSON `imza` alanıyla geçersiz kılınır).
6. **Ek:** kalın-altı çizili `Ek:` + liste. Gövde↔Ek atıf uyumu scriptçe denetlenir. Aylık raporda
   Ek-1 Veri Deseni, ardından dayanak/sözleşme/YK kararı ekleri.

Dil ve biçim: resmî, kısa, edilgen-nötr; Şirket "Şirketimiz", Kurum "Kurumunuz". Tutarlar Türk
formatı `1.250.000,00 TL`. Tarih GG.AA.YYYY. Kararda olmayan yükümlülük/süre/yorum Kurum kararı
gibi yazılmaz.

## Kapsam ve tutarlılık kontrolleri

**Aylık rapor (rapor_olustur.py otomatik denetler; bilinçli ol):**
- "100 bin TL üzeri" = J > 100.000,00; eşik dahil değil. Ana faaliyet kalemleri (komisyon, hasar,
  reasürans, maaş, vergi, yasal yükümlülükler) rapora alınmaz; gri alanlarda (aidat, avukatlık,
  sponsorluk) varsayılan DAHİL et ve teslim notunda belirt.
- Hukuk dosyası ödemeleri 01 sayfasına konmaz — günlük bildirim (tür 5) kapsamındadır; üst yazıda
  ayrıca atıf/beyan isteniyorsa `notlar` alanına cümle ekle.
- 02 sayfası kapsam hücre notu: **2026 mali tablolarında giderleştirilmiş TÜM sözleşmeler**
  (yalnız yeni imzalananlar değil — her raporda güncel kümülatif durum). Sözleşme sürüyorsa
  N (Sonlanma) boş. Hizmet konusu **Sigortacılık Destek Hizmet Yönetmeliği md. 4**
  sınıflandırmasıyla yazılır.
- 03 sayfası: "ilgili ay içerisinde alınan" kararlar — ay dışı karar tarihi uyarı verir.
- 500.000 TL aşan sözleşme: YK kararı ve Kurum onayı (tür 4) bağlantısını denetle; onaysızsa
  uyarı ver ve tür 4 yazısını hatırlat.
- Yıl/Ay sütunları GÖNDERİM dönemini taşır; ödemenin ait olduğu ay K sütununa (Ödeme Ayı).
- Ek dosyaları veri klasörlerinde gerçekten var mı kontrol edilir; yoksa uyarı — gönderim
  öncesi tamamlanmalı.
- Şirket kodu (D sütunu) bilinmiyorsa kullanıcıya sor; boş bırakılırsa uyarı verilir.

**Sözleşme onayı (tür 4):** üç şart birlikte (ana faaliyet dışı + yükümlülük + toplam bedel
500.000 TL üzerinde). Toplam bedel esastır (tüm süre/taksitler). Onay imzadan önce; taahhüt
cümlesi zorunlu. Zaten imzalıysa dürüstçe kur ve kullanıcıyı uyar.

**Hukuk (tür 5/6):** tablo toplamı ile bakiye cümlesi tutarlı: `kalan = 20.000.000 − önceki −
günün`. Tür 6 önce gönderilmiş olmalı. Gerçek kişi kimlik/IBAN'ı yazma (KVKK).

**Düzeltme (tür 8):** hatalı yazının tarih/Ref'i İlgi (b)'ye; eski→yeni karşılaştırma net.

**Genel:** rakamları kaynaktan iki kez kontrol et; uydurma yerine `[●...]`; kişisel veri/ticari
sır içeren eklerde maskeleme öner; yazı e-imza ile KEP/EBYS kanalından gider (seddk@hs01.kep.tr);
Kurum süre vermişse tarihleri karşılaştır, gecikmede tek cümle gerekçe öner.

## Script Kullanımı

```bash
# AYLIK RAPOR (xlsx + üst yazı birlikte)
python scripts/rapor_olustur.py girdi.json --ustyazi
python scripts/rapor_olustur.py girdi.json --validate-only      # sadece doğrula
python scripts/rapor_olustur.py --tara --veri-dizini "..."     # klasör envanteri
python scripts/rapor_olustur.py --ornek-girdi iskelet.json       # boş girdi şeması

# OLAY BAZLI YAZI (tür 4-8)
python scripts/yazi_olustur.py girdi.json cikti.docx [--validate-only]

# HER İKİSİNİN SONRASI
python scripts/dogrula.py cikti.docx     # sayfa yerleşimi + yer tutucu denetimi (pymupdf)
```

Girdi şemaları script docstring'lerinde; örnekler `references/ornekler/`.

**Zorunlu doğrulamalar** (hata = üretim durur): zorunlu alanlar, GG.AA.YYYY, tablo sütun uyumu,
tablo toplam satırı aritmetiği, İlgi (x)/Ek-N atıf uyumu, imza yapısı, ödeme eşiği (aylık
rapor), rapor sütun/satır bütünlüğü.

**Uyarılar** (üretim sürer, gözden geçir): tek ilgiye harfli atıf, Türk para formatı dışı hücre,
eksik kapanış, karar tarihinden önceki yazı tarihi, eksik ek dosyası, ay dışı YK kararı,
500.000 TL aşan sözleşmede onay durumu, üst yazı tarihinin 15'i aşması.

## Teslim notu (yanıtta kısaca)

- Ay/yıl; Veri Deseni satır sayıları (ödeme/sözleşme/YK) ve ödeme toplamı
- Ref/tarih veya yer tutucular; eksiklerin listesi
- Ek listesi; kapsam uyarıları (gri alan kalemi, onaysız 500 bin sözleşme, hukuk günlük ayrımı)
- `dogrula.py` sonucu (sayfa sayısı dahil) ve dosyaların tam yolları

## Sık Karşılaşılan Hatalar

1. **Kararda/üst yazıda olmayan hüküm üretmek.** Eşik ve periyot yalnızca kararın ve raporlama
   üst yazısının metninden.
2. **Formatı bozmak.** Veri Deseni sütun adları/sırası değiştirilirse, tip satırı raporda
   bırakılırsa veya Notlar bölümü silinirse Kurum reddeder.
3. **Toplam satırını elle yazıp aritmetik hataya düşmek.** Script denetler; uyuşmazlıkta üretim
   durur.
4. **Hukuk ödemesini 01 sayfasına karıştırmak.** Hukuk günlük rejimdedir (tür 5).
5. **02 sayfasını yalnız yeni sözleşmelerle doldurmak.** Kapsam: 2026'da giderleştirilmiş TÜM
   sözleşmeler (kümülatif).
6. **Yer tutucuyu uydurma veriyle doldurmak.** `[●...]` sarı vurguda kalır; gönderim öncesi liste
   boş olmalı.
7. **15'i kaçırmak.** Rapor ayı takip edilir; üst yazı tarihi son teslim gününü aşıyorsa script
   uyarır.
8. **Yanlış imzacı.** Aylık rapor: iç denetim müdürü + iç sistemlerden sorumlu YK üyesi (rapor
   üst yazısının şartı). Diğer yazılar: varsayılan ikili.

## Doğrulama Listesi

- [ ] Raporlama rejimi doğru: aylık rapor mu, olay bazlı yazı mı
- [ ] İlgi (a) = 1937 Karar künyesi; aylık raporda İlgi (b) = E-51048673-010.05-5045010
- [ ] Eşikler doğru: 100 bin TL aşan ödeme / 500.000 TL üzerinde sözleşme / 20.000.000 TL bloke
- [ ] Veri Deseni sütun yapısı bire bir; tip satırı ve Notlar raporda yok
- [ ] Yıl/Ay = gönderim dönemi; Ödeme Ayı ayrı; N sütunu devam eden sözleşmede boş
- [ ] Tablo toplamları ve satır sayıları aritmetik olarak doğru (script doğruladı)
- [ ] Gövde↔İlgi ve gövde↔Ek atıf uyumu tamam (script doğruladı)
- [ ] İmzacılar doğru: aylık raporda iç denetim müdürü + iç sistemlerden sorumlu YK üyesi
- [ ] `rapor_olustur.py`/`yazi_olustur.py` hatasız + `dogrula.py` HATA'sız; yer tutucular listelendi
- [ ] Dosya adları kalıpta; teslim notu verildi

## Dosya Haritası

- `references/veri_deseni.md` — Kurum Veri Deseni'nin sütun sütun talimatı + hücre notları + klasör sözleşmesi
- `references/sablonlar.md` — olay bazlı yazı şablonları (tür 4-9 + birleşik), dil ve üslup notları
- `references/ornekler/` — 10 hazır girdi JSON'u (olay bazlı türler)
- `assets/sirket_bilgileri.json` — muhatap, karar + raporlama künyeleri, imzacılar, şirket kodu
- `scripts/rapor_olustur.py` — aylık rapor üretici: Veri Deseni xlsx + üst yazı docx
- `scripts/yazi_olustur.py` — olay bazlı yazı üretici + kapsamlı girdi doğrulama
- `scripts/dogrula.py` — sayfa yerleşimi/yer tutucu denetleyicisi (pymupdf)