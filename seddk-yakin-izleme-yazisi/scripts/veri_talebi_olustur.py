#!/usr/bin/env python3
"""Birimlerden dış hizmet alımı verisi toplamak için FORM Excel'i üretir (v4.1).

Kaynak kurallar:
- SEDDK E-51048673-010.05-5045010 sayılı "Kuruma Yapılacak Raporlamalar Hk." üst yazısı eki
  Veri Deseni ('02-Dış Hizmet' sayfası) — sütun yapısı bire bir korunur (A-S), ardından
  İç Denetim süreç sütunları (T-W) eklenir.
- Sigortacılık Destek Hizmetleri Hakkında Yönetmelik MADDE 4 — E sütunu 'Hizmet Konusu'
  sınıflandırması (a-j bentleri + md.4/2 bilgi sistemleri + md.1/2 kapsam dışı grubu).
  SEDDK hücre notu: "Sigortacılık Destek Hizmet Hakkında Yönetmeliğin 4'üncü maddesine göre
  sınıflandırılmalıdır. Söz konusu sınıflandırmaya uymayan konular da ayrıca olarak
  belirtilmelidir."

Ürettiği dosya: 3 sheet
  1. 'AÇIKLAMA'  — neden istiyoruz, kapsam, md.4 sınıflandırma listesi, kılavuz, tanımlar.
  2. 'DIŞ HİZMET BEYAN FORMU' — doldurulacak tablo; dropdown'lar mümkün olduğunca listbox.
  3. 'LISTE' — dropdown kaynak listeleri (kullanıcı değiştirmeyecek).

Dropdown'lar (listbox):
  E  Hizmet Konusu (md.4 sınıfı, tam etiketli)        <- LISTE sayfası
  H  Hizmet Bedelinin Hesaplanma Şekli                <- LISTE sayfası
  R  YK Kararı Alındı mı? (Evet/Hayır)                <- satır içi
  T  YK Kararı Alınacak mı? (Evet/Hayır)              <- satır içi
  V  Kurum Onayı Durumu                              <- satır içi
  B  Ay (1-12 tam sayı doğrulama), A Yıl (tam sayı)
  K/L/M tarih, N/O/P tutar doğrulaması — uyarı stilinde (örn. 'Süresiz' gibi girdileri
     engellememek için sert değil), E/H/R/T/V dropdown'ları sert.

Kullanım:
  python veri_talebi_olustur.py <çıktı.xlsx> [--sayi 30]
"""
import argparse
import os
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------- Yönetmelik md.4 sınıflandırma listesi ----------------
# (kod, dropdown kısa etiketi, tam bent metni)
MD4_LISTE = [
    ("a",  "MD4(a) — Teknik inceleme ve kontrol hizmetleri",
           "Poliçe tanzimi ile tazminat tedvir ve ödenmesine ilişkin süreçlerde, sigorta "
           "eksperliği işinden ayrı olmak kaydıyla gerçekleştirilen teknik inceleme ve kontrol "
           "hizmetleri"),
    ("b",  "MD4(b) — Risk azaltma / zarar azaltma hizmetleri",
           "Hasar öncesi risk azaltmaya ve hasar sonrası zarar azaltmaya yönelik hizmetler"),
    ("c",  "MD4(c) — Hasar ihbarı, dosya açma ve tamamlama",
           "Hasar ihbarı alma, dosya açma ve tamamlama hizmetleri"),
    ("ç",  "MD4(ç) — Onarım ve bakım hizmetleri",
           "Onarım ve bakım hizmetleri"),
    ("d",  "MD4(d) — Yedek parça tedarik ve kontrol",
           "Yedek parça tedarik ve kontrol hizmetleri"),
    ("e",  "MD4(e) — Yardım (asistans) hizmetleri",
           "Yardım (asistans) hizmetleri"),
    ("g",  "MD4(g) — Tedavi ve yardım hizmetleri",
           "Tedavi ve yardım hizmetleri"),
    ("ğ",  "MD4(ğ) — Çağrı merkezi hizmetleri",
           "Çağrı merkezi hizmetleri"),
    ("h",  "MD4(h) — Sovtaj yönetimi hizmetleri",
           "Sovtaj yönetimi hizmetleri"),
    ("ı",  "MD4(ı) — Rücu takip hizmetleri",
           "Rücu takip hizmetleri"),
    ("i",  "MD4(i) — Arşiv yönetimi hizmetleri",
           "Arşiv yönetimi hizmetleri"),
    ("j",  "MD4(j) — Ürün ve tarife hazırlama hizmetleri",
           "Ürün ve tarife hazırlama hizmetleri"),
    ("BS", "MD4(BS) — Bilgi sistemleri (md.4/2)",
           "Bilgi sistemleri; sigortacılık faaliyet ve yükümlülükleri bakımından yönetim, içerik "
           "tasarımı, erişim, kontrol, denetim, güncelleme, bilgi/rapor alma gibi fonksiyonlarda "
           "karar alma gücünün ve sorumluluğun şirkette olması şartıyla destek hizmeti alımına "
           "konu edilebilir (md.4/2)"),
]
# NOT: f bendi (tıbbi danışmanlık) Danıştay 10. Daire 17/5/2021 E.:2016/1892; K.:2021/2273
# kararıyla iptal — dropdown'da yok; AÇIKLAMA'da not düşülür.

KAPSAM_DISI = [
    ("KD1", "KD1 — Personel istihdamına ilişkin hizmet alımı (kapsam dışı)",
            "Başka bir kurum bünyesinde istihdam edilmekle birlikte şirkette geçici veya sürekli "
            "olarak çalıştırılan personele ilişkin hizmet alımları (Yön. md.1/2-a)"),
    ("KD2", "KD2 — Avukatlık / her türlü danışmanlık (kapsam dışı)",
            "Avukatlık hizmetleri, vergi ve hukuk danışmanlığı dâhil her türlü danışmanlık "
            "(Yön. md.1/2-b)"),
    ("KD3", "KD3 — Reklam faaliyetleri (kapsam dışı)",
            "Reklam faaliyetleri (Yön. md.1/2-c)"),
]

HESAP_SEKLI = ["Belirli Bedel", "Üretime Bağlı", "Hasara Bağlı", "Kullanıma Bağlı",
               "Faydalanmaya Bağlı", "Diğer (I sütununda açıklayınız)"]
EVET_HAYIR = ["Evet", "Hayır"]
ONAY_DURUM = ["Kurum onayı gerekmiyor", "Kurum onayı alındı",
              "Kurum onayı bekleniyor", "Bilinmiyor / araştırılıyor"]

# Form sütunları: (başlık, genişlik, hücre notu, zorunlu mu)
SUTUNLAR = [
    ("Yıl",                               8,  "Bilginin gönderim dönemi yılı. Önceden dolu — değiştirmeyin.", True),
    ("Ay",                                6,  "Gönderim dönemi ayı (1-12). Önceden dolu — değiştirmeyin.", True),
    ("Şirket Adı",                        22, "Önceden dolu — değiştirmeyin.", True),
    ("Şirket Kodu",                       10, "SEDDK şirket kodu — İç Denetim dolduracak, boş bırakın.", False),
    ("Hizmet Konusu *",                   44, "Sigortacılık Destek Hizmetleri Hakkında Yönetmeliğin 4'üncü "
                                              "maddesine göre sınıflandırınız (SEDDK hücre notu). Listeden "
                                              "seçin. Hiçbiri uymuyorsa 'KD… (kapsam dışı)' grubundan seçin — "
                                              "sınıflandırmaya uymayan konular da raporda ayrıca belirtilir. "
                                              "Birden fazla bent kapsamındaysa en baskın olanı seçip özet "
                                              "içerikte belirtin.", True),
    ("Alınan Hizmetin Özet İçeriği *",    42, "Hizmetin kapsamı — 1-3 cümle özet.", True),
    ("Hizmetin Alınma Amacı *",           38, "Şirket hangi ihtiyacı gideriyor? KD (kapsam dışı) seçtiyseniz "
                                              "burada kısaca açıklayın.", True),
    ("Hizmet Bedelinin Hesaplanma Şekli *", 26, "Sözleşme imzalanırken belli ise 'Belirli Bedel'; üretim/hasar/"
                                              "kullanım/faydalanma sonucuna göre değişiyorsa ilgili 'Bağlı' "
                                              "seçeneği (SEDDK hücre notu).", True),
    ("Birim Maliyeti",                    15, "Varsa birim maliyet (dosya başı, araç başı, kişi başı vb.). "
                                              "Yoksa boş bırakın (SEDDK hücre notu: niteliğine göre varsa).", False),
    ("Hizmet Sağlayıcısının Adı *",       30, "Tam unvan.", True),
    ("Sözleşme Başlangıç Tarihi *",      16, "Sözleşmede belirtilen başlama tarihi (GG.AA.YYYY).", True),
    ("Sözleşme Bitiş Tarihi *",           16, "Sözleşmede belirtilen bitiş tarihi (GG.AA.YYYY). Süresiz ise "
                                              "'Süresiz' yazın.", True),
    ("Sözleşme Sonlanma Tarihi",          16, "Sözleşme DEVAM EDİYORSA BOŞ BIRAKIN (SEDDK hücre notu). "
                                              "Feshedildi/bittiyse sonlandığı tarih.", False),
    ("Sözleşmenin Toplam Bedeli (TL) *",  18, "Sözleşmede belirtilen bedel veya tahmini bedel (SEDDK hücre "
                                              "notu). Süreklilik gösteren hizmetlerde yıllık tahmini bedel. "
                                              "KDV hariçse (KDV hariç) notu ekleyin.", True),
    ("Tahakkuk Eden Hizmet Bedeli (TL) *", 16, "2026 yılı içinde bu hizmete tahakkuk eden bedel. Yoksa 0.", True),
    ("Ödenen Kısım (TL) *",               14, "2026 yılı içinde ödenen tutar. Yoksa 0.", True),
    ("Ödemenin Kaydedildiği Hesap Kodu",  14, "Muhasebe hesap kodu (örn. 654). Bilmiyorsanız boş bırakın — "
                                              "Mali İşler'ten alınacak.", False),
    ("YK Kararı Alındı mı? *",            18, "Bu hizmet/sözleşme için Yönetim Kurulu kararı alındıysa 'Evet' "
                                              "— tarih/sayı biliniyorsa S sütununa yazın.", True),
    ("YK Karar Tarihi / Sayısı",          18, "Varsa YK kararının tarihi ve sayısı (örn. 12.09.2026 / 987).", False),
    ("YK Kararı Alınacak mı?",            18, "Karar henüz alınmadıysa ve alınması gerekiyorsa 'Evet' — "
                                              "planlanan tarih/sayı bilgisini S sütununa ekleyin.", False),
    ("Sözleşme + YK Kararı Dosya Adları", 32, "Sözleşme ve YK kararı PDF'lerinin adı (örn. "
                                              "2026-09-01_CagriMerkezi_Sozlesme.pdf). PDF'leri formu "
                                              "gönderirken e-postaya ekleyin — İç Denetim SOZLESMELER/YKK "
                                              "klasörlerine arşivleyecek.", False),
    ("Kurum Onayı Durumu",                24, "TOPLAM bedeli 500.000 TL'yi aşan ana faaliyet dışı sözleşmeler "
                                              "SEDDK onayına tabidir (1937 Kararı md.2); onay imzadan ÖNCE "
                                              "alınır. Eşiğin altındaysa 'gerekmiyor' seçin. Kırmızı hücre "
                                              "= eşiği aştınız, İç Denetim'i bilgilendirin.", False),
    ("Birim / İletişim *",                26, "Beyanı veren birim + sorumlu kişinin ad-soyadı, e-posta, telefon.", True),
]

# form sütun başlığı -> SEDDK Veri Deseni '02-Dış Hizmet' sütun başlığı (None = form only)
FORM2SEDDK = {
    "Yıl": "Yıl", "Ay": "Ay", "Şirket Adı": "Şirket Adı", "Şirket Kodu": "Şirket Kodu",
    "Hizmet Konusu *": "Hizmet Konusu", "Alınan Hizmetin Özet İçeriği *": "Alınan Hizmetin Özet İçeriği",
    "Hizmetin Alınma Amacı *": "Hizmetin Alınma Amacı",
    "Hizmet Bedelinin Hesaplanma Şekli *": "Hizmet Bedelinin Hesaplanma Şekli",
    "Birim Maliyeti": "Birim Maliyeti", "Hizmet Sağlayıcısının Adı *": "Hizmet Sağlayıcısının Adı",
    "Sözleşme Başlangıç Tarihi *": "Sözleşme Başlangıç Tarihi",
    "Sözleşme Bitiş Tarihi *": "Sözleşme Bitiş Tarihi",
    "Sözleşme Sonlanma Tarihi": "Sözleşme Sonlanma Tarihi",
    "Sözleşmenin Toplam Bedeli (TL) *": "Sözleşmenin Toplam Bedeli",
    "Tahakkuk Eden Hizmet Bedeli (TL) *": "Tahakkuk Eden Hizmet Bedeli",
    "Ödenen Kısım (TL) *": "Ödenen Kısım",
    "Ödemenin Kaydedildiği Hesap Kodu": "Ödemenin Kaydedildiği Hesap Kodu",
    "YK Kararı Alındı mı? *": "YK Kararı Alındı mı?",
}


def build(out_path, n_rows, yil=2026, ay=10):
    wb = Workbook()
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    gri = PatternFill("solid", fgColor="E7E6E6")
    mavi = PatternFill("solid", fgColor="DDEBF7")
    sari = PatternFill("solid", fgColor="FFF2CC")
    kirmizi_fill = PatternFill("solid", fgColor="FFC7CE")
    kirmizi_font = Font(color="9C0006")

    # ================= Sheet 3 önce oluşturulur (kaynak listeler) =================
    wsL = wb.create_sheet("LISTE")
    wsL.sheet_view.showGridLines = False
    wsL.cell(row=1, column=1, value="Bu sayfa formdaki açılır listelerin kaynağıdır — LÜTFEN DEĞİŞTİRMEYİN. "
                                    "Sınıflandırmalar SİGORTACILIK DESTEK HİZMETLERİ HAKKINDA YÖNETMELİK md.4'ten "
                                    "alınmıştır.").font = Font(bold=True, size=10, color="C00000")
    wsL.column_dimensions["A"].width = 55
    wsL.column_dimensions["B"].width = 100
    wsL.column_dimensions["D"].width = 34
    wsL.cell(row=2, column=1, value="HİZMET KONUSU (md.4 sınıflandırması)").font = Font(bold=True, size=10)
    wsL.cell(row=2, column=2, value="YÖNETMELİK MADDE 4 / 1-2 BENT METNİ").font = Font(bold=True, size=10)
    wsL.cell(row=2, column=4, value="BEDEL HESAPLAMA ŞEKLİ").font = Font(bold=True, size=10)
    for i, (kod, kisa, tam) in enumerate(MD4_LISTE + KAPSAM_DISI):
        wsL.cell(row=3 + i, column=1, value=kisa)
        wsL.cell(row=3 + i, column=2, value=tam).alignment = Alignment(wrap_text=True, vertical="top")
    n_md4 = len(MD4_LISTE) + len(KAPSAM_DISI)
    for i, hs in enumerate(HESAP_SEKLI):
        wsL.cell(row=3 + i, column=4, value=hs)
    E_RANGE = f"LISTE!$A$3:$A${2 + n_md4}"
    H_RANGE = f"LISTE!$D$3:$D${2 + len(HESAP_SEKLI)}"

    # ================= Sheet 1: AÇIKLAMA =================
    ws = wb.active
    ws.title = "AÇIKLAMA"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 108
    ws.column_dimensions["C"].width = 3

    def P(row, text, bold=False, size=11, color=None, fill=None, italic=False):
        c = ws.cell(row=row, column=2, value=text)
        c.font = Font(bold=bold, size=size, italic=italic, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        ws.row_dimensions[row].height = max(15, 13 * (1 + len(text) // 110))
        return c

    r = 1
    P(r, "DIŞ HİZMET ALIMI BEYAN FORMU — BİLGİLENDİRME", bold=True, size=14); r += 1
    P(r, "DOGA SİGORTA A.Ş. — İç Denetim Müdürlüğü veri talebi | SEDDK yakın izleme kapsamı",
      size=10, color="808080"); r += 2

    P(r, "1. NEDEN İSTİYORUZ?", bold=True, size=12, fill=gri); r += 1
    P(r, "SEDDK, Şirketimizi E-97354901-045.01-5033535 sayılı üst yazısı ile 23.09.2026 tarih ve 1937 sayılı Kurul "
         "Kararı uyarınca YAKIN İZLEME kapsamına almıştır. Karar uyarınca; ana sigortacılık faaliyeti kapsamındaki "
         "ödemeler hariç 100 bin TL ÜZERİ ödemeler, yönetim kurulu kararları ve DIŞ HİZMET ALIMLARINA ilişkin "
         "bilgilerin Kuruma bildirilmesi yükümlülüğü bulunmaktadır."); r += 1
    P(r, "Kurum ayrıca E-51048673-010.05-5045010 sayılı \"Kuruma Yapılacak Raporlamalar Hk.\" üst yazısı ile bu "
         "bilgilerin, yazı ekindeki VERİ DESENİ adlı formatta, biten aya ilişkin raporun HER AYIN 15'İNE KADAR, "
         "iç denetim müdürü ve iç sistemlerden sorumlu yönetim kurulu üyesi imzalı yazı ile sunulmasını istemiştir. "
         "Raporun dış hizmet sayfasında, 2026 YILI MALİ TABLOLARINDA GİDERLEŞTİRİLMİŞ TÜM SÖZLEŞMELERİN "
         "(ödenen veya daha önce ödenmekle birlikte 2026 mali tablolarına yansıtılan giderlere konu sözleşmeler) "
         "her ay güncel durumlarıyla yer alması gerekmektedir."); r += 1
    P(r, "Bu form, söz konusu aylık raporun \"02-Dış Hizmet\" sayfasının hazırlanması amacıyla birimlerimizden "
         "toplanacak bilgilerin standart biçimidir. Beyan ettiğiniz her sözleşme/hizmet SEDDK'ya raporlanacak ve "
         "denetim izine konu olacaktır. Eksik veya hatalı beyan hem Şirketimize hem beyan sahibine idari "
         "yükümlülük doğurur.", fill=sari); r += 1
    P(r, "Ayrıca 1937 sayılı Karar gereği, sigortacılık ana faaliyeti dışında yükümlülük doğuran ve tutarı "
         "500.000 TL'Yİ AŞAN sözleşmeler İMZALANMADAN ÖNCE SEDDK onayına tabidir. Formda bu eşiği aşan "
         "sözleşmeleriniz kırmızı renkle işaretlenir — İç Denetim'e ayrıca bilgi verin; imza öncesi onay "
         "süreci başlatılacaktır.", fill=sari); r += 2

    P(r, "2. KAPSAM — HANGİ SÖZLEŞMELER BEYAN EDİLECEK?", bold=True, size=12, fill=gri); r += 1
    P(r, "Aşağıdaki koşulları birlikte sağlayan her dış hizmet alımı için bir satır doldurun:"); r += 1
    P(r, "    1) Sözleşme 2026 yılında yürürlükte (başlamış) VEYA 2026 mali tablolarına gider yansıtılmış "
         "(tahakkuk etmiş) — bitmiş olsa bile 2026'da gider yansıyanlar DAHİL;"); r += 1
    P(r, "    2) Hizmet Şirketin ana sigortacılık faaliyeti dışında kalan bir alanda (destek hizmeti, "
         "danışmanlık, bilişim, arşiv, çağrı merkezi, temizlik, güvenlik vb.);"); r += 1
    P(r, "    3) Alım bir sözleşmeye/taahüde dayanıyor (hizmet sözleşmesi, çerçeve sözleşme, ihale, satın alma "
         "emri, sözleşme ekleri/uzatmalar dahil)."); r += 1
    P(r, "BEYAN DIŞI: komisyon, hasar ödemesi, reasürans ödemeleri, maaş-ücret (bordro) ödemeleri, vergi ve "
         "diğer yasal yükümlülükler — bunlar dış hizmet alımı değildir, bu formda yer almaz."); r += 2

    P(r, "3. HİZMET KONUSU NASIL SINIFLANDIRILIR? (E sütunu — açılır liste)", bold=True, size=12, fill=gri); r += 1
    P(r, "SEDDK'nın Veri Deseni hücre notu: \"Sigortacılık Destek Hizmet Hakkında Yönetmeliğin 4'üncü maddesine "
         "göre sınıflandırılmalıdır. Söz konusu sınıflandırmaya uymayan konular da ayrıca olarak "
         "belirtilmelidir.\""); r += 1
    P(r, "Buna göre E sütunundaki açılır listeden aşağıdaki sınıflardan birini seçin (tam bent metinleri "
         "LISTE sayfasındadır):"); r += 1
    for kod, kisa, tam in MD4_LISTE:
        P(r, f"    {kisa}", size=10); r += 1
    P(r, "Yönetmeliğin md.1/2'sinde KAPSAM DIŞI tutulan (yani md.4 listesine girmeyen) hizmetler de — SEDDK "
         "hücre notu gereği raporda AYRICA belirtilmek üzere — formda beyan edilir:"); r += 1
    for kod, kisa, tam in KAPSAM_DISI:
        P(r, f"    {kisa}", size=10); r += 1
    P(r, "NOT 1: Md.4'ün (f) bendi (hasar tedvir uygulamalarında tıbbi danışmanlık hizmetleri) Danıştay Onuncu "
         "Dairesinin 17/5/2021 tarihli ve E.:2016/1892; K.:2021/2273 sayılı kararı ile iptal edilmiştir; listede "
         "yoktur. Bu tür hizmetleri 'MD4(g) — Tedavi ve yardım hizmetleri' altında değerlendirin.",
      size=9, italic=True); r += 1
    P(r, "NOT 2: Birden fazla bent kapsamına giren hizmetlerde en baskın (ana) sınıfı seçin; diğer bentleri "
         "'Alınan Hizmetin Özet İçeriği' sütununda belirtin.", size=9, italic=True); r += 1
    P(r, "NOT 3: Hiçbir bent uymuyorsa uygun 'KD… (kapsam dışı)' seçeneğini işaretleyin ve 'Hizmetin Alınma "
         "Amacı' sütununda kısaca açıklayın.", size=9, italic=True); r += 2

    P(r, "4. DOLDURMA KILAVUZU", bold=True, size=12, fill=gri); r += 1
    P(r, "• Her sözleşme/hizmet alımı için bir satır doldurun. İlk satır örnek (sarı) — kopyalayıp çoğaltın, "
         "sonra içeriğini kendi verinizle değiştirin."); r += 1
    P(r, "• Açılır listeli hücrelerde (E, H, R, T, V sütunları) hücreye tıklayıp açılan oktan seçim yapın; "
         "serbest metin yazmayın."); r += 1
    P(r, "• Tarihler GG.AA.YYYY (örn. 05.10.2026); tutarlar TL (örn. 150.000,00). Bitiş tarihi süresiz "
         "sözleşmelerde 'Süresiz' yazabilirsiniz."); r += 1
    P(r, "• 'Sözleşme Sonlanma Tarihi' (M): sözleşme halen yürürlükteyse BOŞ bırakın."); r += 1
    P(r, "• Başlığında * olan sütunlar zorunludur; diğerleri elinizde mevcutsa doldurun."); r += 1
    P(r, "• Emin olmadığınız hücreyi BOŞ bırakın ve formu gönderirken e-postada belirtin — TAHMİN YAZMAYIN.",
      fill=sari); r += 1
    P(r, "• SEDDK formatında olmayan ek sütunlar (S, T, U, V, W) İç Denetimin süreç yönetimi içindir: YK karar "
         "detayı, beklenen YK kararı, PDF dosya adları, Kurum onayı durumu, birim/iletişim. Doldurulması "
         "beklenmektedir."); r += 1
    P(r, "• Satır yetmezse son satırı kopyalayıp alta ekleyin — tablo biçimi korunur."); r += 2

    P(r, "5. SÜREÇ VE SORUMLULUK", bold=True, size=12, fill=gri); r += 1
    P(r, "• Formu dolduran birim, formu birim yöneticisinin bilgisi dâhilinde İç Denetim Müdürlüğüne "
         "(Serkan KOÇ) e-posta ile iletir; sözleşme ve YK kararı PDF'lerini de ekler."); r += 1
    P(r, "• İç Denetim formları birleştirir, SEDDK Veri Deseni'ne aktarır, eksikleri tamamlar, belgeleri "
         "arşiv klasörlerine koyar ve aylık raporu Kuruma sunar."); r += 1
    P(r, "• Sonradan hata/eksik fark edilirse formu düzeltip yeniden gönderin — Kuruma giden raporda düzeltme "
         "yazısı gerekir, gecikme risklidir."); r += 1
    P(r, "• İlk dolum için hedef süre bu e-postada belirtilmiştir. Sonraki aylarda aynı form güncel durumla "
         "(tahakkuk/ödenen sütunları değişmiş olabilir) yeniden istenecektir."); r += 2

    P(r, "6. TANIMLAR", bold=True, size=12, fill=gri); r += 1
    P(r, "• Destek Hizmeti Sağlayıcısı: Sigortacılık Kanunu ve BES Kanunu kapsamında faaliyet gösteren "
         "şirketlere, faaliyet alanlarıyla ilgili konularda yardımcı veya tamamlayıcı nitelikte hizmet veren "
         "kişi veya kuruluş (Yön. md.3-c)."); r += 1
    P(r, "• Ana faaliyet: poliçe üretimi, prim/hasar/reasürans yönetimi, acentelik-brokerlik ve mevzuatın "
         "şirkete yüklediği diğer çekirdek işler. Destek hizmeti bunların yardımcı/tamamlayıcısıdır."); r += 1
    P(r, "• Bilgi Merkezi: Sigorta Bilgi ve Gözetim Merkezi (SBM). Destek hizmeti alımları için Yön. md.5/1 "
         "risk-fayda-maliyet raporu SBM'ye iletilir; md.5/9 uyarınca yıllık değerlendirme her yıl Mart sonuna "
         "kadar SBM'ye ve yönetim kuruluna sunulur."); r += 1
    P(r, "• KEP: Kayıtlı Elektronik Posta — SEDDK resmî yazışmalarında kullanılır (seddk@hs01.kep.tr)."); r += 2

    P(r, "7. GİZLİLİK VE KİŞİSEL VERİ", bold=True, size=12, fill=gri); r += 1
    P(r, "Bu form şirket içi denetim ve SEDDK raporlaması amaçlıdır (KVKK md.5/2-ç, -e kapsamında). Forma "
         "gerçek kişilere ait T.C. kimlik numarası, IBAN, doğum tarihi gibi veri YAZMAYIN; tüzel kişi unvanları "
         "veya görev unvanıyla ad-soyad yeterlidir. Form içeriği gizlilik dereceli şirket bilgisidir — "
         "şirket dışıyla paylaşmayın."); r += 1

    # ================= Sheet 2: FORM =================
    ws2 = wb.create_sheet("DIŞ HİZMET BEYAN FORMU")
    nc = len(SUTUNLAR)

    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    t = ws2.cell(row=1, column=1,
                 value=f"DIŞ HİZMET ALIMI BEYAN FORMU — {yil}/{ay:02d} dönemi | "
                       f"DOGA SİGORTA A.Ş. (ayrıntılı bilgi için 'AÇIKLAMA' sayfasına bakınız)")
    t.font = Font(bold=True, size=12, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="1F4E78")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    for j, (baslik, gen, not_, zorunlu) in enumerate(SUTUNLAR, start=1):
        c = ws2.cell(row=2, column=j, value=baslik)
        c.font = Font(bold=True, size=9, color="1F4E78")
        c.fill = mavi
        c.border = border
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws2.column_dimensions[get_column_letter(j)].width = gen
        cm = Comment(not_, "İç Denetim")
        cm.width, cm.height = 340, 190
        c.comment = cm
    ws2.row_dimensions[2].height = 44
    ws2.freeze_panes = "C3"

    first_data, last_data = 3, 2 + n_rows

    def dv_list_formula(formula, cell_range, prompt, title, strict=True):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                            showDropDown=False)
        if strict:
            dv.error = "Lütfen açılır listeden bir değer seçin."
            dv.errorTitle = "Geçersiz değer"
        dv.prompt = prompt
        dv.promptTitle = title
        ws2.add_data_validation(dv)
        dv.add(cell_range)
        return dv

    # E: Hizmet Konusu (md.4) — LISTE sayfasından
    dv_list_formula(E_RANGE, f"E{first_data}:E{last_data}",
                    "Sigortacılık Destek Hizmetleri Yön. md.4 sınıflandırması (tam liste LISTE ve "
                    "AÇIKLAMA sayfalarında). Hiçbiri uymuyorsa 'KD… (kapsam dışı)' seçin.",
                    "Hizmet Konusu *")
    # H: Hesaplama şekli — LISTE sayfasından
    dv_list_formula(H_RANGE, f"H{first_data}:H{last_data}",
                    "Bedel sözleşmede belli mi, yoksa üretim/hasar/kullanım/faydalanmaya bağlı mı?",
                    "Hesaplanma Şekli *")
    # R / T: Evet-Hayır (satır içi)
    dv_list_formula('"Evet,Hayır"', f"R{first_data}:R{last_data}",
                    "Bu hizmet/sözleşme için YK kararı alındıysa Evet; tarih/sayı S sütununa.",
                    "YK Kararı Alındı mı? *")
    dv_list_formula('"Evet,Hayır"', f"T{first_data}:T{last_data}",
                    "Karar henüz alınmadıysa ve alınması gerekiyorsa Evet.",
                    "YK Kararı Alınacak mı?")
    # V: Kurum onayı durumu
    dv_list_formula(f'"{",".join(ONAY_DURUM)}"', f"V{first_data}:V{last_data}",
                    "500.000 TL üzeri ana faaliyet dışı sözleşmeler imza ÖNCESİ SEDDK onayına tabidir.",
                    "Kurum Onayı Durumu")

    # A yıl, B ay tam sayı
    dvy = DataValidation(type="whole", operator="between", formula1="2020", formula2="2035",
                         allow_blank=False, errorStyle="stop")
    dvy.error = "Yıl 2020-2035 arasında olmalı."
    ws2.add_data_validation(dvy); dvy.add(f"A{first_data}:A{last_data}")
    dvm = DataValidation(type="whole", operator="between", formula1="1", formula2="12",
                         allow_blank=False, errorStyle="stop")
    dvm.error = "Ay 1-12 arasında olmalı."
    ws2.add_data_validation(dvm); dvm.add(f"B{first_data}:B{last_data}")

    # K/L/M tarih — uyarı stilinde (Süresiz gibi girdilere izin)
    for cl in ("K", "L", "M"):
        dv = DataValidation(type="date", operator="between",
                            formula1="DATE(2020,1,1)", formula2="DATE(2035,12,31)",
                            allow_blank=True, errorStyle="warning")
        dv.error = ("Tarih GG.AA.YYYY biçiminde olmalı (örn. 05.10.2026). 'Süresiz' gibi metin "
                    "girişi kabul edilir ama lütfen kontrol edin.")
        dv.errorTitle = "Tarih biçimi uyarısı"
        ws2.add_data_validation(dv); dv.add(f"{cl}{first_data}:{cl}{last_data}")

    # N/O/P tutar — uyarı stilinde
    for cl in ("N", "O", "P"):
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                            allow_blank=True, errorStyle="warning")
        dv.error = "Tutar TL ve 0 veya daha büyük olmalı (örn. 150000 veya 150000,50)."
        dv.errorTitle = "Tutar uyarısı"
        ws2.add_data_validation(dv); dv.add(f"{cl}{first_data}:{cl}{last_data}")

    # N > 500.000 koşullu biçim (kırmızı)
    from openpyxl.formatting.rule import CellIsRule
    ws2.conditional_formatting.add(
        f"N{first_data}:N{last_data}",
        CellIsRule(operator="greaterThan", formula=["500000"],
                   fill=kirmizi_fill, font=kirmizi_font, stopIfTrue=False))

    # ---- satırlar ----
    # örnek satır (3)
    ornek = [
        yil, ay, "DOGA SİGORTA ANONİM ŞİRKETİ", "",
        "MD4(ğ) — Çağrı merkezi hizmetleri",
        "7/24 çağrı merkezi ilk seviye destek hattı işletimi (ÖRNEK SATIR — kopyalayıp kullanın, "
        "sonra bu satırı silin)",
        "Poliçe ve hasar süreçlerinde sigortalı iletişiminin kesintisiz yürütülmesi (örnek)",
        "Belirli Bedel", "—", "ÖRNEK Çağrı Hizmetleri A.Ş.",
        "01.09.2026", "31.12.2026", None, "150.000,00", "37.500,00", "37.500,00",
        "654", "Evet", "12.09.2026 / 987", "Hayır",
        "2026-09-01_CagriMerkezi_Sozlesme.pdf / 2026-09-12_YK_987.pdf",
        "Kurum onayı gerekmiyor",
        "Operasyon Birimi / Ayşe YILMAZ / a.yilmaz@… / 0212 …",
    ]
    for j, v in enumerate(ornek, start=1):
        c = ws2.cell(row=first_data, column=j, value=v)
        c.fill = sari
        c.border = border
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        c.font = Font(italic=True, size=9, color="7F7F7F")
    ws2.cell(row=first_data, column=1).number_format = "0"
    ws2.cell(row=first_data, column=2).number_format = "0"
    ws2.row_dimensions[first_data].height = 58

    # boş satırlar
    for rr in range(first_data + 1, last_data + 1):
        for j in range(1, nc + 1):
            c = ws2.cell(row=rr, column=j)
            c.border = border
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.font = Font(size=9)
        ws2.cell(row=rr, column=1, value=yil).font = Font(size=9)
        ws2.cell(row=rr, column=2, value=ay).font = Font(size=9)
        ws2.cell(row=rr, column=3, value="DOGA SİGORTA ANONİM ŞİRKETİ").font = Font(size=9)
        ws2.row_dimensions[rr].height = 26

    info = ws2.cell(row=last_data + 2, column=1,
                    value="Not: Satır yetmezse son satırı kopyalayıp alta ekleyin. Zorunlu alanlar (*) "
                          "boş kalmasın. Sorular için İç Denetim Müdürlüğü.")
    info.font = Font(italic=True, size=9, color="808080")

    ws2.sheet_view.zoomScale = 80
    # sheet sırası: AÇIKLAMA, FORM, LISTE
    wb.move_sheet("LISTE", offset=1)
    wb.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description="Birim dış hizmet beyan formu üretir (v4.1)")
    ap.add_argument("cikti", help="Çıktı .xlsx")
    ap.add_argument("--sayi", type=int, default=30, help="Boş satır sayısı (varsayılan 30)")
    ap.add_argument("--yil", type=int, default=2026, help="Dönem yılı")
    ap.add_argument("--ay", type=int, default=10, help="Dönem ayı (1-12)")
    a = ap.parse_args()
    out = build(a.cikti, a.sayi, a.yil, a.ay)
    print(f"Form üretildi: {out}")
    print("Sheet'ler: AÇIKLAMA | DIŞ HİZMET BEYAN FORMU | LISTE")
    print(f"Satırlar: 1 örnek (sarı) + {a.sayi} boş (yıl/ay/şirket ön dolu)")
    print("Dropdown'lar: E=md.4 sınıfı (LISTE kaynaklı), H=bedel şekli (LISTE), R/T=Evet/Hayır, "
          f"V=Kurum onayı durumu; A/B tam sayı; K/L/M tarih (uyarı), N/O/P tutar (uyarı); "
          f"N>500.000 kırmızı koşullu biçim.")


if __name__ == "__main__":
    main()