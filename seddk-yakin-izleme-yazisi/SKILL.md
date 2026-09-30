---
name: seddk-yakin-izleme-yazisi
description: Use when the user needs any official letter to SEDDK under the 23.09.2026 date and 1937-numbered Board decision putting DOGA Sigorta A.Ş. under close surveillance (yakın izleme) — payment notifications above 100.000 TL outside core insurance activity, board decision notifications, outsourced service notifications, contract approval requests above 500.000 TL, daily law-file payment notifications, the 20.000.000 TL blocked-amount account notification, internet banking transfer-authority removal notification, correction/additional-info letters, or periodic status summaries. Also triggers on "SEDDK'ya yazı", "Kuruma bildirim", "yakın izleme yazısı", "Kurum onayı yazısı", "hukuk ödemesi bildirimi", "düzeltme yazısı". Produces a Konu/İlgi/Açıklama/Ek structured, dual-signed Word (.docx) letter in the company's house format, with pre-flight scope and arithmetic validation.
version: 3.0.0
author: DOGA Sigorta İç Denetim / Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [seddk, yakin-izleme, resmi-yazi, doga-sigorta, bildirim, kurum-onayi]
    related_skills: [audit-report-generation]
---

# SEDDK Yakın İzleme Yazısı Hazırlama (v3)

## Genel Bakış

SEDDK, Şirketi E-97354901-045.01-5033535 sayılı üst yazısıyla 23.09.2026 tarih ve 1937 sayılı Kurul Kararı uyarınca yakın izleme kapsamına almıştır. Bu skill, karardan doğan bildirim ve onay yazılarını **tek bir tutarlı sürecle** hazırlar:

türü belirle → kapsam kontrolü yap → eksikleri tek seferde sor → JSON'u kur → scripti çalıştır → doğrulayıcıyı çalıştır → teslim et.

Resmî yazışmada tek bir kusur (yanlış İlgi, eksik ek, aritmetik hata, taşan tablo) Kurumun geri dönüşüne yol açar. Bu yüzden bu skill yalnızca şablon değil, **yazıdan önce ve sonra zorunlu kontrol disiplini** tanımlar. Veri uydurma yok; eksikler `[●açıklama]` yer tutucusu olarak işaretlenir ve script bunları sarı vurgular.

## Kararın birebir metni (tek doğruluk kaynağı)

Tüm eşik ve kapsam kuralları bu metne dayanır; başka yerden eşik okuma, kararın koyduğu hükümleri yorumla genişletme. Karar üst yazısı (E-97354901-045.01-5033535) dört karardan oluşur:

| # | Karar (birebir) | Doğan yazı türü |
|---|---|---|
| 1 | "Şirketinizin yakın izleme kapsamına alınmasına **ve** ana sigortacılık faaliyeti kapsamındaki ödemeler hariç (komisyon ödemeleri, hasar ödemeleri, reasürans ödemeleri, maaş ödemeleri, vergi ve diğer yasal yükümlülükler gibi) **100 bin TL üzeri ödemeler**, **yönetim kurulu kararları** ve **dış hizmet alımlarına** ilişkin bilgileri yakın izleme kapsamında Kurumumuza göndermesine" | 1 Ödeme, 2 YK kararı, 3 Dış hizmet |
| 2 | "Sigortacılık ana faaliyeti dışında yükümlülük doğuran ve tutarı **500.000 TL üzerinde** kalan tüm sözleşmelerin Kurumumuz **onayına** tabi tutulmasına" | 4 Sözleşme onay talebi |
| 3 | "Şirketinizin internet bankacılığı aracılığıyla para transferi (Şirketinizin **kendi hesapları arasındaki işlemler hariç**) yetkilerinin meblağ, banka, finansal kuruluş ve konudan bağımsız olarak **tümüyle kaldırılmasına**" | 7 Uygulama bildirimi (karar zorunlu kılmıyor, belgeleme amaçlı) |
| 4 | "Hukuk dosyaları kapsamında yapılan ödemelerin **günlük** olarak Kurumumuza bildirilmesi ve serbest bırakılacak tutarın **sadece hukuk dosyalarında kullanılması** kaydıyla, Şirketiniz tarafından belirlenecek hesapta tutulacak **20.000.000 TL'lik bloke tutarının serbest bırakılmasına" | 5 Günlük bildirim, 6 Hesap bildirimi |

Karar üst yazısının tam künyesi, muhatabı, imzacıları: `assets/sirket_bilgileri.json` → `karar_kunyesi`. Gövdede karar tarih/sayısına atıf her zaman bu künede bire bir yazılan üst yazı ve Karar ile yapılır.

Ek türler (kararın doğrudan emretmediği ama pratikte ihtiyaç duyulan): **8 Düzeltme/ek bilgi**, **9 Periyodik durum özeti**.

## Yazı türleri ve şablonlar

Her türün Konu kalıbı, İlgi düzeni, gövde iskeleti, toplanacak bilgiler ve tipik ekleri `references/sablonlar.md` içinde. **Yazıya başlamadan önce ilgili türü oradan oku** — burada tekrar edilmez. Hazır girdi örnekleri `references/ornekler/` altında (9 türün tamamı + birleşik yazı):

`01_odeme`, `02_yk_karari`, `03_dis_hizmet`, `04_sozlesme_onay`, `05_hukuk_gunluk`, `06_hesap_bildirimi`, `07_internet_bankaciligi`, `08_duzeltme`, `09_periyodik_ozet`, `10_birlesik`.

## İş Akışı

1. **Türü belirle.** Kullanıcının saydığı işlemleri yukarıdaki türlere eşle. Birden fazla işlem varsa `sablonlar.md` bölüm 10'daki birleştirme kuralını uygula: **onay talepleri (4), hukuk günlük bildirimi (5) ve düzeltme (8) her zaman ayrı yazıdır**; 1-2-3 aynı dönemdeyse numaralı alt başlıklarla birleştirilebilir. Tamamlandı kriteri: her işlem tam olarak bir türe eşlenmiş olmalı.
2. **Kapsam kontrolü yap** (aşağıdaki "Kapsam ve tutarlılık kontrolleri"). Kontrol, yazının hiç gerekmediğini gösterirse (ör. tamamı maaş/vergi) bunu açıkça söyle ve yazı hazırlama.
3. **Eksikleri tek seferde sor** — yalnızca yazıyı anlamlı kılanları (tutar, karşı taraf, tarih, dönem). Kullanıcının verdiği bilgiyi bir daha sorma. İkincil eksikler için `[●açıklama]` yer tutucusu bırak; script sarı vurgular. Tamamlandı kriteri: zorunlu alanların hepsi ya dolu ya bilinçli yer tutucu.
4. **Girdi JSON'unu kur.** `references/ornekler/` içinden ilgili türü kopyalayıp uyarla. Dosya adı: `SEDDK_<TurAdi>_<GGAAYYYY>.json` (örn. `SEDDK_Sozlesme_Onay_Talebi_30092026.json`).
5. **Üret ve doğrula:**
   ```bash
   python scripts/yazi_olustur.py girdi.json SEDDK_<TurAdi>_<GGAAYYYY>.docx
   python scripts/dogrula.py SEDDK_<TurAdi>_<GGAAYYYY>.docx   # sayfa yerleşimi + yer tutucu denetimi
   ```
   Script hata verirse düzelt ve yeniden çalıştır; uyarıları okuyup bilinçli karar ver. Tamamlandı kriteri: `yazi_olustur.py` hatasız, `dogrula.py` HATA'sız dönmüş olmalı. Not: `dogrula.py` sayfa ölçümünü pymupdf yerleşimiyle yapar; Word `keepNext` zincirini uyguladığından gerçek sayfa sayısı bir-iki sayfa daha az olabilir — sayfa uyarısı bilgi amaçlıdır, HATA sayılmaz.
6. **Teslim.** Yanıtta "Teslim notu" bölümünü ver (aşağıda).

Python ortamı (bu makinede): ağır kütüphaneler (python-docx, pymupdf) yalnızca sistem Python 3.12'de kurulu; scriptleri onunla çalıştır:
`C:/Users/DS_COM70/AppData/Local/Programs/Python/Python312/python.exe -u scripts/yazi_olustur.py ...` (`-u` çıktı tamponlamasını kapatır). Başka ortamda python-docx kuruluysa normal `python` yeter.

## Zorunlu yazı yapısı (ev formatı)

Biçim, Şirketin Kuruma gönderdiği gerçek yazılardan (Pert/Sovtaj iç denetim üst yazısı, İç Denetim Envanteri eksik evrak yazısı, sözleşme onay talebi örneği) alınmıştır; script bunu üretir:

1. **Muhatap (solda, kalın, 3 satır):** SİGORTACILIK VE ÖZEL EMEKLİLİK / DÜZENLEME VE DENETLEME KURUMU / İSTANBUL. Sağda kalın **Tarih: GG.AA.YYYY** ve **Ref: YYYY/NNN** (şirket giden evrak numarası; kullanıcı vermediyse `[●YYYY/NNN]`). Antetli kâğıt kullanıldığı için sayfa üstünde antet için boşluk bırakılır, antet basılmaz.
2. **Konu:** tek cümlelik, işlemi ve (varsa) dönemi/karşı tarafı adlandıran başlık.
3. **İlgi:** en az bir satır. (a) her zaman Kurumun E-97354901-045.01-5033535 sayılı üst yazısı ve 1937 sayılı Karar (`sirket_bilgileri.json` → `karar_kunyesi.ana_ilgi`). Sonraki harfler ilgili önceki yazışmalar. Birden fazla ilgi a), b), c) diye **harflenir**; gövdede "İlgi (a) yazınız" diye atıf yapılır. **Tek ilgi harflenmez** ve gövdedeki atıf "İlgi yazınız ile ..." olur — script bu uyumu denetler.
4. **Açıklama (gövde):** iki yana yaslı, ilk satır girintili, 1,15 satır aralıklı. İlk paragrafta yazının **amacı** açıkça yazar ("... bilgi verilmesi amaçlanmaktadır", "... onayınız talep edilmektedir"); sonra ayrıntı/tablo/beyan; en sonda girintisiz tek satır kapanış: "Bilgilerinize arz ederiz." / "Gereğini ve onayınızı arz ederiz." / "Bilgilerinize ve gereğine arz ederiz." Script kapanış cümlesini denetler.
5. **Saygılarımızla, / DOGA SİGORTA A.Ş.** (ortalı) ve altında **yan yana iki imza**: sol **Serkan KOÇ – Müdür**, sağ **Kuntay BAYDAR – Genel Müdür Yardımcısı**. Farklı imzacı istenirse JSON `imza` alanıyla geçersiz kılınır (geçmişte Coşkun GÖLPINAR – Genel Müdür, Fehmi ÖZBALKAN – GMY imzalamıştır). İmza bloğu bölünmeye karşı tek parçadır.
6. **Ek:** varsa imzanın altında kalın-altı çizili `Ek:` başlığı ve `- ...` listesi. Gövdede atıf yapılan her ek listede, listedeki her ek gövdede geçmelidir — script "Ek-2" atfı ile liste uzunluğunu karşılaştırır. Sayfa sayısı biliniyorsa yaz.

Dil ve biçim: resmî, kısa, edilgen-nötr ("arz ederiz", "sunulmuştur"); Şirket "Şirketimiz", Kurum "Kurumunuz". Tutarlar Türk formatı `1.250.000,00 TL`; önemli tutarlarda yazıyla teyit ("1.250.000,00 TL (birmilyonikiyüzellibin Türk Lirası)"). Tarih GG.AA.YYYY. Kurum adı ilk geçtiği yerde tam yazılır. Kararda olmayan bir yükümlülüğü, süreyi veya yorumu Kurum kararıymış gibi yazma.

## Kapsam ve tutarlılık kontrolleri (yazıdan önce)

**Ödeme bildirimi (tür 1)**
- Ana faaliyet kapsamı hariçliği kararın kendi örneklendirmesiyle: komisyon (acente/broker), hasar, reasürans, maaş, vergi ve diğer yasal yükümlülükler **hariç**. Gri kalan kalemlerde (sektör kuruluşu aidatı, avukatlık ücreti, bağış/sponsorluk) **varsayılan olarak bildirime dahil et** ve bunu teslim notunda belirt.
- "100 bin TL üzeri" = 100.000 TL'yi **aşan**; eşit tutar kapsam dışı. Döviz ödemede TL karşılığını, kurunu ve kur tarihini yaz. KDV dahil/hariç açıkça yazılmalı; emin değilsen sor (muhafazakâr varsayım: KDV dahil).
- Bölünmüş ödeme: aynı alıcıya aynı işten parçalanmış ödemeler birlikte eşiği aşıyorsa kullanıcıyı uyar ve toplamın bildirilmesini öner.
- Hukuk dosyası ödemeleri tür 5'in günlük rejimine tabidir; dönemsel ödeme tablosuna alma, gövdede atıf yap.
- Bildirimin ödemeden önce mi sonra mı yapılacağı kararda açık değil; kullanıcı belirtmediyse **"ödeme sonrası" varsay** ve teslim notunda işaretle.
- Md. 3 gereği internet bankacılığından transfer yapılamayacağı için ödeme kanalı mutlaka gerçek yöntemle yazılmalı (havale/EFT/çek...).

**Sözleşme onayı (tür 4)**
- Üç şart birlikte: ana faaliyet dışı + yükümlülük doğuruyor + toplam bedel 500.000 TL üzerinde. **Toplam bedel** esastır (tüm süre, tüm taksitler); yıllık dilim eşiğin altında olsa bile toplam aşıyorsa onay gerekir. "Üzerinde" = aşan.
- Ek protokol, süre uzatımı, yenileme, bedel artışı ve aynı tedarikçiye bağlı çerçeve/sipariş sözleşmeleri eşiği aşan toplamlarında uyar.
- Onay **imzadan önce** alınır; yazıda "onay alınmadan imzalanmayacak ve yürürlüğe girmeyecektir" taahhüdü bulunur. Sözleşme zaten imzalıysa yazıyı dürüstçe kur (gerekçeli bilgi + onay talebi) ve kullanıcıyı uyar.
- Onay sonrası yapılacak 100.000 TL aşan ödemeler ayrıca tür 1 ile bildirilir — gövdede bu atıf yer alsın.

**Yönetim kurulu kararı (tür 2)**: Karar "yönetim kurulu kararları" diyor, ayrım yok; kullanıcı aksini teyit etmedikçe **tüm YK kararları bildirilir** varsay ve bunu hatırlat. Karar tarih/sayısı ve özeti ek suretiyle birebir uyumlu olmalı.

**Dış hizmet alımı (tür 3)**: Sağlayıcı, kapsam, süre, bedel, kritik faaliyet niteliği, risk değerlendirmesi eksiksiz olmalı. Bedel 500.000 TL'yi aşıyorsa ve ana faaliyet dışıysa tür 4 de hazırlanmalı.

**Hukuk günlük bildirimi (tür 5/6)**: Tablo toplamı ile bakiye cümlesi aritmetik olarak tutarlı olmalı: `kalan = 20.000.000 − önceki kümülatif − günün toplamı`. Script tablo toplamını denetler; bakiye cümlesini sen kur. Tür 6 hesap bildirimi ilk günlük bildirimden önce gönderilmiş olmalı; gönderilmemişse hatırlat. Ödeme olmayan günlerde "ödeme yoktur" yazısı kararda öngörülmemiş; belirsizse yalnız ödeme olan günleri hazırla ve notunu işaretle. Gerçek kişi kimlik/IBAN'ı gövdeye yazma (KVKK, asgari veri).

**Düzeltme (tür 8)**: Hatalı yazının tarih ve Ref'ini İlgi (b)'ye doğru taşı; eski→yeni karşılaştırmayı net yaz; diğer bilgilerin aynı olduğunu belirt.

**Genel**
- Rakamları kaynaktan iki kez kontrol et; script toplamları hesaplar ama girdi doğru olmalı.
- Bilinmeyen tutar/tarih/sayı için uydurma değil `[●...]` bırak.
- Ekte kişisel veri/ticari sır varsa maskeleme öner.
- Yazı e-imza ile KEP/EBYS kanalından gider (Kurum KEP: seddk@hs01.kep.tr); ekleri PDF hazırlamayı ve dosya adlarını ek sırasıyla vermeyi öner.
- Kurum süre vermişse tarihleri karşılaştır; gecikme varsa tek cümle gerekçe öner.
- Gelen yazıya cevaptaysa Kurum yazısının tarih/sayısını İlgi'ye taşı.

## Script Kullanımı

```bash
python scripts/yazi_olustur.py girdi.json SEDDK_<TurAdi>_<GGAAYYYY>.docx   # üret
python scripts/yazi_olustur.py girdi.json cikti.docx --validate-only      # sadece doğrula
python scripts/dogrula.py SEDDK_<TurAdi>_<GGAAYYYY>.docx                  # sayfa yerleşimi
```

**Girdi şeması** `scripts/yazi_olustur.py` başındaki docstring'de; örnekler `references/ornekler/` dizininde.

**Üretim öncesi zorunlu doğrulamalar** (hata = üretim durur): zorunlu alanlar, GG.AA.YYYY formatı, tablo sütun uyumu, **tablo toplam satırının satırların toplamıyla aritmetik eşitliği**, İlgi (x) atıflarının ilgi listesiyle uyumu, Ek-N atıflarının ek listesiyle uyumu, imza yapısı.

**Uyarılar** (üretim sürer ama gözden geçir): tek ilgi varken "İlgi (a)" atfı, çok ilgi varken harfsiz atıf, Türk para formatı dışı hücre, eksik kapanış cümlesi, karar tarihinden önceki yazı tarihi, dolu ek listesine gövde atıfsızlığı.

**Sayfa düzeni:** satır aralığı 1,15; kapanış→Saygı→imza→Ek bloğu `keep_with_next` ile bağlı; tablo satırları bölünmez; başlık satırı sayfa sonunda tekrarlanır. `scripts/dogrula.py` sayfa sayısını, tek kalan imza/ek bloğunu, yer tutucuları ve kenar taşmasını denetler.

## Teslim notu (yanıtta kısaca)

- Yazı türü ve Konu; Ref ve tarih (veya yer tutucu)
- Tablo/toplam özeti; eksik yer tutucuların listesi
- Ek listesi; kapsam uyarıları (gri alan kalemi, imzalanmış sözleşme, eksik hesap bildirimi, ödeme öncesi/sonrası varsayımı)
- `dogrula.py` sonucu (sayfa sayısı dahil) ve dosyanın yolu

## Sık Karşılaşılan Hatalar

1. **Kararda olmayan hüküm üretmek.** "Karar günlük bildirim diyor" değil "Karar hukuk dosyası ödemelerinin günlük bildirilmesine karar vermiştir" — eşik ve kapsam yalnızca kararın metninden.
2. **Tek ilgiyi "İlgi (a)" diye atfetmek.** Resmî yazışmada tek ilgi harflenmez; gövdede "İlgi yazınız ile ...". Çok ilgi varsa harf şart. Script uyarır.
3. **Toplam satırını elle yazıp aritmetik hataya düşmek.** Toplamı hesapla, JSON'a yaz; script yine denetler ve uyuşmazlıkta üretimi durdurur.
4. **Ek atfı ile ek listesinin uyumsuz kalması.** Gövdede "Ek-2'de sunulmuştur" varsa listede en az 2 madde olmalı; script denetler.
5. **Yer tutucuyu uydurma veriyle doldurmak.** `[●...]` sarı vurguda kalır; gönderim öncesi `dogrula.py` çıktısındaki yer tutucu listesi boş olmalı.
6. **Sayfa düzenini kontrol etmeden teslim etmek.** 3+ sayfa ya da tek başına kalan imza bloğu = `dogrula.py` uyarısı; metni kısalt ya da tabloyu sığdır.
7. **Hukuk ödemesini dönemsel tabloya karıştırmak.** Hukuk dosyası ödemeleri günlük rejime tabidir (tür 5); tür 1 tablosuna konmaz.

## Doğrulama Listesi

- [ ] Tür kararla eşleşiyor; kapsam kontrolü yapıldı, gri alanlar işaretlendi
- [ ] İlgi (a) = Kurum üst yazısı + 1937 sayılı Karar künyesi (karar_kunyesi'den bire bir)
- [ ] Eşikler doğru: 100 bin TL **aşan** ödeme / 500.000 TL **üzerinde** sözleşme / 20.000.000 TL bloke
- [ ] Tablo toplamları aritmetik olarak doğru (script doğruladı)
- [ ] Gövde↔İlgi ve gövde↔Ek atıf uyumu tamam (script doğruladı)
- [ ] Kapanış cümlesi türle uyumlu; imzacılar doğru (varsayılan veya geçersiz kılınmış)
- [ ] `yazi_olustur.py` hatasız + `dogrula.py` HATA'sız; yer tutucular bilinçli ve listelendi
- [ ] Dosya adı `SEDDK_<TurAdi>_<GGAAYYYY>.docx` kalıbında; teslim notu verildi

## Dosya Haritası

- `references/sablonlar.md` — 9 tür + birleşik yazı şablonları, dil ve üslup notları
- `references/ornekler/` — 10 hazır girdi JSON'u (her tür + birleşik)
- `assets/sirket_bilgileri.json` — muhatap, karar künyesi (ana_ilgi), varsayılan imzacılar, üst boşluk
- `scripts/yazi_olustur.py` — docx üretici + kapsamlı girdi doğrulama
- `scripts/dogrula.py` — sayfa yerleşimi/yer tutucu denetleyicisi (pymupdf)