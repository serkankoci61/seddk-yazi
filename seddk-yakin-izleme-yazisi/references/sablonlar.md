# Yazı Türleri, Şablonlar ve Zorunlu Bilgiler

Her tür için: Konu kalıbı, İlgi düzeni, gövde iskeleti, toplanacak bilgiler, tipik ekler.
`[●...]` yer tutucuları kullanıcıdan gelmeyen bilgiler içindir; asla uydurma veri yazma.

Eşikler (1937 sayılı Karar'dan, başka yerden değil):

| Kavram | Eşik | Anlamı |
|---|---|---|
| Ödeme bildirimi | 100 bin TL **üzeri** | 100.000 TL'yi **aşan**; eşik tutarın kendisi kapsam dışı |
| Sözleşme onayı | 500.000 TL **üzerinde** | Toplam bedeli (tüm süre + taksitler) 500.000 TL'yi aşan |
| Bloke tutar | 20.000.000 TL | Yalnızca hukuk dosyası ödemelerinde kullanılacak serbest tutar |
| Bildirim periyodu (hukuk) | Günlük | Ödeme yapılan her gün |

Ortak giriş paragrafı (her yazıda ilk paragraf; tek ilgi harflenmez):
> İlgi yazınız ile Sigortacılık ve Özel Emeklilik Düzenleme ve Denetleme Kurulu'nun 23.09.2026 tarih ve 1937 sayılı Kararı uyarınca Şirketimizin yakın izleme kapsamına alındığı bildirilmiştir.

Çok ilgi varsa:
> İlgi (a) yazınız ile ... yakın izleme kapsamına alındığı bildirilmiştir.

Ortak kapanış cümleleri (girintisiz, son satır):
- Bildirim/bilgi yazıları: **"Bilgilerinize arz ederiz."**
- Onay/talep yazıları: **"Gereğini ve onayınızı arz ederiz."**
- Düzeltme/ek bilgi: **"Bilgilerinize ve gereğine arz ederiz."**

---

## 1. Ödeme bildirimi (ana faaliyet dışı, 100.000 TL aşan)

**Kapsam:** Karar md. 1 — ana sigortacılık faaliyeti kapsamındaki ödemeler hariç (komisyon, hasar, reasürans, maaş, vergi ve diğer yasal yükümlülükler) ve tek tek 100.000 TL'yi aşan ödemeler.

**Konu:** `Yakın İzleme Kapsamında Ödeme Bilgilendirmesi ([●başlangıç] – [●bitiş] Dönemi)` (tek ödeme: `... ([●alıcı unvanı])`)

**İlgi:** (a) Kurum üst yazısı + Karar künyesi.

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar kapsamında, ana sigortacılık faaliyeti dışında kalan ve 100.000 TL'yi aşan ödemelere ilişkin bilgiler, [●dönem] dönemi için aşağıdaki tabloda sunulmuştur."
2. Tablo: Sıra | Ödeme Tarihi | Alıcı Unvanı | Ödeme Nedeni (fatura/sözleşme no) | Tutar (TL) | Ödeme Kanalı/Banka | Dayanak (YK kararı / Kurum onayı)
3. Toplam satırı (script aritmetik doğrular).
4. Beyan: "Tabloda yer alan ödemeler komisyon, hasar, reasürans, maaş, vergi ve diğer yasal yükümlülük niteliğinde değildir; hukuk dosyaları kapsamındaki ödemeler ise ayrıca günlük olarak bildirilmektedir."
5. Kapanış: "Bilgilerinize arz ederiz."

**Toplanacak bilgiler:** ödeme tarihi, alıcı, neden, tutar (KDV dahil/hariç açık), para birimi ve kur, ödeme kanalı (internet bankacılığı hariç — md. 3), dayanak belge, ilgili sözleşme/YK kararı/Kurum onayı.
**Tipik ekler:** ödeme dekontları, fatura suretleri, (varsa) sözleşme ve onay yazısı, ödeme listesi (Excel).

---

## 2. Yönetim kurulu kararı bildirimi

**Konu:** `Yakın İzleme Kapsamında Yönetim Kurulu Kararının Bildirilmesi ([●GG.AA.YYYY] Tarih ve [●no] Sayılı Karar)`

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar uyarınca Şirketimiz Yönetim Kurulu'nun [●tarih] tarih ve [●no] sayılı kararının sureti ilişikte sunulmuştur."
2. Kısa karar özeti (konu, karar, varsa tutar/yükümlülük, uygulama tarihi). Özet kararın içeriğini değiştirmez; bire bir ek surete atıf yapılır.
3. Kararın Kurum onayına tabi işlem içerip içermediği (ör. 500.000 TL aşan sözleşme) — içeriyorsa ayrı onay yazısına atıf.
4. Kapanış: "Bilgilerinize arz ederiz."

**Not:** Karar "yönetim kurulu kararları" diyor, ayrım yok; kullanıcı aksini teyit etmedikçe tüm YK kararları bildirilir. Bu varsayımı teslim notunda hatırlat.
**Tipik ekler:** YK kararı (imzalı/e-imzalı suret), varsa dayanak sunum/rapor.

---

## 3. Dış hizmet alımı bildirimi

**Konu:** `Yakın İzleme Kapsamında Dış Hizmet Alımı Bilgilendirmesi ([●hizmet sağlayıcı unvanı])`

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar kapsamında, aşağıda bilgileri yer alan dış hizmet alımı hakkında Kurumunuza bilgi verilmesi amaçlanmaktadır."
2. Bilgi tablosu (Başlık | Bilgi): sağlayıcı unvanı/vergi no/adres; hizmetin konusu ve kapsamı; Şirket faaliyetlerindeki yeri (kritik mi); sözleşme tarihi ve süresi; bedel (KDV durumu, ödeme planı); alım gerekçesi; yetkili onay (YK/yetki); risk/uyum değerlendirmesi özeti; ilişkili taraf durumu (varsa).
3. Bedel 500.000 TL'yi aşıyorsa ve ana faaliyet dışıysa: "Sözleşme bedeli 500.000 TL'yi aştığından Kurum onayı için ayrıca [●tarih/sayı] yazımızla başvurulmuştur" (tür 4 de hazırlanır; onay imzadan önce).
4. Kapanış: "Bilgilerinize arz ederiz."

**Tipik ekler:** sözleşme/teklif, SLA dokümanı, risk değerlendirme formu, YK/yetki kararı.

---

## 4. Sözleşme onay talebi (ana faaliyet dışı, 500.000 TL üzerinde)

**Kapsam:** Karar md. 2 — sigortacılık ana faaliyeti dışında, yükümlülük doğuran ve toplam tutarı 500.000 TL üzerinde kalan tüm sözleşmeler Kurum onayına tabi. Onay imzadan önce alınır.

**Konu:** `500.000 TL Üzerindeki Sözleşmeye İlişkin Kurum Onayı Talebi ([●sözleşme konusu] – [●karşı taraf unvanı])`

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar uyarınca, sigortacılık ana faaliyeti dışında yükümlülük doğuran ve tutarı 500.000 TL'yi aşan sözleşmeler Kurumunuzun onayına tabi tutulmuştur. Bu çerçevede, aşağıda özeti sunulan sözleşme için onayınız talep edilmektedir."
2. Özet tablo (Başlık | Bilgi): karşı taraf unvan/VKN, sözleşme konusu, süre, **toplam bedel** (KDV durumu), ödeme planı, gerekçe, YK/yetki kararı, risk değerlendirmesi.
3. Taahhüt: "Sözleşme, Kurumunuzun onayı alınmadan imzalanmayacak ve yürürlüğe girmeyecektir."
4. Onay sonrası ödeme bildirimi atfı: "Sözleşme kapsamında yapılacak ve 100.000 TL'yi aşan ödemeler ayrıca yakın izleme kapsamında bildirilecektir."
5. Kapanış: "Gereğini ve onayınızı arz ederiz."

**Toplanacak bilgiler:** karşı taraf unvan/VKN, sözleşme konusu, taslak metin, toplam bedel, süre, ödeme planı, YK kararı tarih/no, risk değerlendirmesi.
**Tipik ekler:** sözleşme taslağı, teklif karşılaştırması, YK kararı, risk/uyum formu.

---

## 5. Hukuk dosyası ödemelerinin günlük bildirimi

**Kapsam:** Karar md. 4 — hukuk dosyası ödemeleri günlük bildirilir; 20.000.000 TL serbest tutar yalnızca bu amaçla kullanılır.

**Konu:** `Hukuk Dosyası Ödemelerine İlişkin Günlük Bildirim ([●GG.AA.YYYY] Tarihi)`

**İlgi:** (a) Kurum üst yazısı + Karar; (b) hesap bildirim yazısı (tür 6, tarih ve Ref).

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar uyarınca [●tarih] tarihinde hukuk dosyaları kapsamında yapılan ödemeler aşağıda tablo halinde sunulmuştur."
2. Tablo: Sıra | Dosya No / Mahkeme | Alıcı (avukat/alacaklı) | Ödeme Nedeni | Tutar (TL) | Kanal
3. Toplam satırı + bakiye cümlesi: "Serbest tutardan gün başı bakiye: [●] TL; günün ödemeleri: [toplam] TL; gün sonu kalan: [hesapla] TL." (kalan = gün başı − günün toplamı; script toplamı denetler, bakiyeyi sen hesapla)
4. Ödeme yoksa (kullanıcı isterse): "Anılan tarihte hukuk dosyası kapsamında ödeme yapılmamıştır."
5. Kapanış: "Bilgilerinize arz ederiz."

**Not:** Gerçek kişi kimlik/IBAN'ı gövdeye yazılmaz (KVKK). Tür 6 önce gönderilmiş olmalı.
**Tipik ekler:** ödeme dekontları, (gerekiyorsa) mahkeme/avukat yazısı.

---

## 6. 20.000.000 TL bloke tutar / hesap bildirimi

**Kapsam:** Karar md. 4 — "Şirketiniz tarafından belirlenecek hesapta" tutulacak bloke tutarın serbest bırakılması; hesap belirlendiğinde ve değiştiğinde bildirilir.

**Konu:** `Hukuk Dosyası Ödemelerine İlişkin Hesap Bildirimi ve Serbest Tutar Kullanım Taahhüdü`

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar uyarınca serbest bırakılacak 20.000.000 TL'lik bloke tutarının tutulacağı hesap aşağıda bildirilmiştir."
2. Tablo: Banka | Şube | Hesap Adı | IBAN | Para Birimi | Tutar. (İBAN şirket hesabına ait; gerçek kişi verisi içermez.)
3. Taahhüt: "Söz konusu tutar yalnızca hukuk dosyaları kapsamındaki ödemelerde kullanılacak, ödemeler günlük olarak Kurumunuza bildirilecektir."
4. Kapanış: "Bilgilerinize arz ederiz."

**Tipik ekler:** bankadan hesap/bloke durumu yazısı (varsa).

---

## 7. İnternet bankacılığı yetki kaldırma uygulaması bildirimi (isteğe bağlı)

**Kapsam:** Karar md. 3 — internet bankacılığıyla para transferi yetkilerinin (kendi hesapları arası hariç) meblağ, banka, finansal kuruluş ve konudan bağımsız tümüyle kaldırılması. Karar ayrı bir bildirim zorunluluğu koymamıştır; uygulamanın belgelenmesi yazışma güvencesi sağlar.

**Konu:** `İnternet Bankacılığı Para Transferi Yetkilerinin Kaldırılması Hk.`

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar'ın 3 üncü maddesi uyarınca internet bankacılığı aracılığıyla para transferi yetkilerinin tümüyle kaldırılmasına ilişkin yapılan işlemler aşağıda özetlenmiştir."
2. Tablo: Banka | Yetki Kaldırma Tarihi | Kendi Hesapları Arası İşlem İzni (varsa) | Ödeme Yöntemi (yeni)
3. Kapanış: "Bilgilerinize arz ederiz."

**Tipik ekler:** bankalardan alınan yetki kaldırma onay yazıları.

---

## 8. Düzeltme / ek bilgi yazısı

**Konu:** `[●önceki yazı konusu] Bildirimine İlişkin Düzeltme` veya `... Ek Bilgi Sunumu`

**İlgi:** (a) Kurum üst yazısı + Karar; (b) düzeltilen önceki yazı (tarih ve Ref).

**Gövde iskeleti:**
1. Ortak giriş + "İlgi (b) yazımızla sunulan [●konu] bildiriminde [●hata/eksiklik] tespit edilmiş olup, doğru bilgiler aşağıda sunulmaktadır."
2. Düzeltme tablosu: Sıra | Alan | Önceki (hatalı) | Düzeltilmiş (sayı sütunları sağa hizalı)
3. Etki cümlesi: "Dönem toplamı buna göre [●doğru toplam] TL olarak düzeltilmiştir. Diğer bilgiler aynıdır."
4. Kapanış: "Bilgilerinize ve gereğine arz ederiz."

**Tipik ekler:** düzeltilmiş tablo/liste, ilgili belgeler.

---

## 9. Periyodik özet / durum raporu (isteğe bağlı)

**Konu:** `Yakın İzleme Kapsamında [●dönem] Dönemi Durum Özeti`

**Gövde iskeleti:**
1. Ortak giriş + "Anılan Karar kapsamında [●dönem] döneminde gerçekleştirilen bildirim ve işlemlerin özeti aşağıda sunulmuştur."
2. Özet tablo: Bildirim Türü | Adet | Toplam Tutar (varsa) | İlgili Yazılarımız (Ref) — toplam satırı scriptçe denetlenir.
3. Önemli gelişmeler (onay bekleyen sözleşmeler, bloke hesabı bakiyesi, YK kararları özeti).
4. Kapanış: "Bilgilerinize arz ederiz."

**Tipik ekler:** dönemsel özet Excel, yazışma listesi.

---

## 10. Birleşik yazı (birden fazla işlem)

- Onay talepleri (tür 4), hukuk günlük bildirimi (tür 5) ve düzeltme (tür 8) **her zaman ayrı yazıdır**.
- Ödeme (1), YK kararı (2) ve dış hizmet (3) aynı döneme aitse **tek yazıda numaralı alt başlıklarla** birleştirilebilir: `{"baslik": "1. Ödeme Bilgileri"}`, `{"baslik": "2. Yönetim Kurulu Kararları"}` vb.
- Birleşik yazının Konusu: `Yakın İzleme Kapsamında Bilgi Sunumu ([●dönem])`; Ekler sırayla numaralanır ve gövdede "Ek-1", "Ek-2" diye atıf alır (script uyumu denetler).
- Kullanıcı ayrı yazı isterse ayrı hazırla.

---

## 11. Aylık yakın izleme raporu (Veri Deseni sunum yazısı)

**Kaynak:** Kurumun E-51048673-010.05-5045010 sayılı "Kuruma Yapılacak Raporlamalar Hk." üst
yazısı. Karar md. 1 kapsamındaki bilgilerin (100 bin TL üzeri ödemeler, YK kararları, dış hizmet
alımları) **aylık** sunumunu düzenler; Veri Deseni'ne (`references/veri_deseni.md`) ve girdi
şemasına (`scripts/rapor_olustur.py` docstring) bakmadan doldurma. `rapor_olustur.py --ustyazi`
bu şablonu otomatik kurar — elle yazma, girdi JSON'unu kur.

**Konu:** `Yakın İzleme Kapsamında <Ay Adı> <Yıl> Dönemine İlişkin Rapor Sunumu`

**İlgi:** (a) Kurum üst yazısı + 1937 sayılı Karar künyesi; (b) E-51048673-010.05-5045010
sayılı raporlama üst yazısı.

**Gövde iskeleti** (script üretir; burada denetim içindir):
1. Ortak giriş (İlgi (a)).
2. İlgi (b) atfıyla raporlama yükümlülüğünün tekrarı: "... işbu yazı ekindeki şekil ve formatta,
   iç denetim müdürü ve iç sistemlerden sorumlu yönetim kurulu üyesi tarafından imzalı yazı ile
   biten aya ilişkin raporun her ayın 15'ine kadar düzenli olarak sunulması hususunda bilgi
   verilmiştir."
3. Dönem cümlesi: "<Ay> <Yıl> dönemine ilişkin rapor ekte sunulmuştur."
4. Ödeme beyanı: sayı + toplam tutar; **yoksa** "100 bin TL üzeri ödeme gerçekleşmemiştir."
5. Dış hizmet beyanı: "2026 yılı mali tablolarında giderleştirilmiş sözleşmeler..."; **yoksa**
   "...rapora konu dış hizmet alımı sözleşmesi bulunmamaktadır."
6. YKK beyanı: sayı; **yoksa** "Dönem içinde yönetim kurulu kararı alınmamıştır."
7. (Opsiyonel) `notlar` alanından gelen ek beyanlar — örn. hukuk ödemelerinin günlük bildirim
   rejimiyle ayrıca takip edildiği atfı.
8. Kapanış: **"Bilgilerinize arz ederiz."**

**İmzacılar:** iç denetim müdürü (Serkan KOÇ) + iç sistemlerden sorumlu yönetim kurulu üyesi
(raporlama üst yazısının şartı — ad kullanıcıdan teyit edilmeli). Genel yazışma imza ikilisi
(Müdür/GMY) burada KULLANILMAZ.

**Ekler:** Ek-1 Veri Deseni (.xlsx); ardından dayanak belgeler, sözleşme PDF'leri, YK kararı
PDF'leri (Veri Deseni satır sırasıyla).

**Süre:** biten ayın izleyen ayının **15'i**. Gecikiyorsa tek cümle gerekçe (örn. "bankadan
alınan dekont yazılarının temin süresi") — savunma değil, bilgi.

**Boş dönem kuralı:** işlem olmasa da rapor gönderilir; sayfalar boş bırakılır (satır açılmaz)
ve gövdede "gerçekleşmemiştir/alınmamıştır" beyanı verilir.

---

## Dil ve üslup notları

- Resmî, kısa, edilgen-nötr: "arz ederiz", "sunulmuştur", "bildirilmiştir".
- Şirket = "Şirketimiz"; Kurum = "Kurumunuz".
- Tutarlar Türk formatı: `1.250.000,00 TL`; önemli tutarlarda yazıyla teyit: `1.250.000,00 TL (birmilyonikiyüzellibin Türk Lirası)`.
- Tarih: GG.AA.YYYY (nokta ayraçlı).
- Kurum adı ilk geçtiği yerde tam; ardından "Kurumunuz".
- Kararın öngörmediği yükümlülük/süre/yorum, Kurum kararı gibi yazılmaz.
- Gövdedeki "ekte sunulmaktadır" / "Ek-1" atıfları ek listesiyle bire bir uyumlu olur (script denetler).