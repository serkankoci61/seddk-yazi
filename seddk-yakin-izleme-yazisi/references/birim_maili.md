# Birim Veri Talebi E-postası (dış hizmet beyan formu)

`[●...]` alanları göndermeden önce doldurulur. Konu satırı ve gövde tek e-postadır; form ve
Yönetmelik PDF'i ekte gider. CC: birim yöneticileri + [●Gerekli görülen yöneticiler].

---

**Konu:** DIŞ HİZMET ALIMI BEYAN FORMU — SEDDK Yakın İzleme Kapsamında Veri Talebi ([●GG.AA.YYYY] son tarihli)

Sayın İlgili,

Şirketimiz, SEDDK'nın 23.09.2026 tarih ve 1937 sayılı Kurul Kararı uyarınca yakın izleme
kapsamına alınmıştır. Karar uyarınca; ana sigortacılık faaliyeti kapsamındaki ödemeler hariç
olmak üzere 100 bin TL üzeri ödemelere, yönetim kurulu kararlarına ve dışarıdan hizmet
alımlarına ilişkin bilgilerin Kuruma düzenli olarak bildirilmesi yükümlülüğü bulunmaktadır.
Kurum ayrıca bu bilgilerin, belirlediği VERİ DESENİ formatında, biten aya ilişkin raporun
HER AYIN 15'İNE KADAR sunulmasını istemiştir.

Bu kapsamda, biriminiz tarafından yürütülen DIŞ HİZMET ALIMINA İLİŞKİN SÖZLEŞMELERİN
bilgilerine ihtiyaç duyulmaktadır. Ekte yer alan **DIŞ HİZMET BEYAN FORMU**'nu doldurmanız
rica olunur. Formda:

- 2026 yılı mali tablolarında giderleştirilmiş (tahakkuk etmiş) veya yürürlükte olan TÜM dış
  hizmet sözleşmelerinizin her biri için bir satır doldurulmalıdır; yalnızca bu yıl imzalanan
  sözleşmeler değil, önceki yıllardan devam eden ve bu yıl gider yansıyan sözleşmeler de
  dahildir.
- "Hizmet Konusu" sütunu, Sigortacılık Destek Hizmetleri Hakkında Yönetmelik'in 4'üncü
  maddesindeki sınıflandırmaya göre AÇILIR LİSTEDEN seçilir; listeye uymayan konular için
  "kapsam dışı" seçenekleri mevcuttir (ayrıntı formun AÇIKLAMA sayfasındadır).
- Komisyon, hasar ödemesi, reasürans, maaş-bordro ve vergi ödemeleri bu formun kapsamı
  DIŞINDADIR; formu doldarken bu kalemleri eklemeyiniz.

Beyan edilecek sözleşme bulunmuyorsa dahi formu "sözleşme yoktur" notu ile iade etmenizi rica
ederiz; aksi halde raporlama eksik kalacaktır.

Formu doldururken dikkat edilmesi gereken noktalar formun **AÇIKLAMA** sayfasında
madde madde açıklanmıştır. Özellikle belirtmek isteriz ki:

- TOPLAM bedeli **500.000 TL'yi aşan** ve ana sigortacılık faaliyeti dışındaki sözleşmeler,
  Kurul Kararı gereği İMZALANMADAN ÖNCE Kurum onayına tabidir. Bu eşiği aşan sözleşmelerinizi
  formda işaretlemeniz halinde İç Denetim sizinle iletişime geçerek onay sürecini
  başlatacaktır.
- Emin olmadığınız alanları boş bırakınız; tahmin yazmayınız. Eksikleri formu iade ederken
  e-postanızda belirtiniz.

**Son teslim tarihi: [●GG.AA.YYYY]** (SEDDK'ya rapor son teslim günü olan [●GG.AA.YYYY]
tarihinden önce birleştirme ve kontroller için yeterli süre bırakılması amacıyla belirlenmiştir.)

Doldurulan formu ve varsa sözleşme/yönetim kurulu kararı PDF'lerini bu e-postaya yanıt olarak
İç Denetim Müdürlüğü'ne iletmenizi rica ederiz. Sorularınız için [●iletişim kişisi/e-posta/telefon]
 üzerinden ulaşabilirsiniz.

Bilgilerinize ve gereğini rica ederiz.

[●Ad SOYAD]
İç Denetim Müdürü
DOGA SİGORTA A.Ş.

**Ekler:**
1. DIŞ HİZMET BEYAN FORMU (Excel — 3 sayfa: AÇIKLAMA, BEYAN FORMU, LISTE)
2. Sigortacılık Destek Hizmetleri Hakkında Yönetmelik (PDF — md.4 sınıflandırması için referans)

---

## Gönderim notları (İç Denetim için)

- Form `scripts/veri_talebi_olustur.py` ile üretilir: dönem yıl/ay parametreleri formun
  yıl/ay ve başlık şeridine önceden yazılır.
- Birden çok birime gönderilirken formun birebir aynı kopyası kullanılır; birim adı formdaki
  "Birim / İletişim" sütunundan zaten toplanır — dosya adına birim eklemek işe yarar
  (örn. `Dis_Hizmet_Beyan_Formu_Operasyon.xlsx`) ama zorunlu değildir.
- Dönen formlar `rapor_olustur.py --formdan <klasör>` ile tek girdi JSON'unda birleştirilir;
  örnek satır otomatik atlanır.
- Son tarih, SEDDK rapor gününden (her ayın 15'i) en az 5 iş günü önce seçilir.