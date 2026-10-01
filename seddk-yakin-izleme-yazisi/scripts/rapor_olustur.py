#!/usr/bin/env python3
"""SEDDK yakın izleme AYLIK RAPORU üretici (v4) — Kurum Veri Deseni formatında xlsx + üst yazı.

Kaynak: SEDDK E-51048673-010.05-5045010 sayılı "Kuruma Yapılacak Raporlamalar Hk." üst yazısı
(30.09.2026) ekindeki Veri Deseni (SEDDK Yakın İzleme Formatı_01092026.xlsx). Üst yazı gereği
biten aya ilişkin rapor, iç denetim müdürü ve iç sistemlerden sorumlu yönetim kurulu üyesi
tarafından imzalı yazı ile HER AYIN 15'İNE KADAR sunulur.

Ürettikleri:
  1) Veri deseni .xlsx — Kurum şablonunun BİRE BİR kopyası üzerine veri satırları eklenir
     (başlık/tip satırları, hücre notları, 'Notlar' bölümü korunur; veri satırları tip
     satırının altına eklenir, ince çerçeveli ve biçimli yazılır).
  2) Üst yazı .docx — aylık rapor sunum yazısı (yazi_olustur.py altyapısıyla; imzacılar:
     iç denetim müdürü + iç sistemlerden sorumlu YK üyesi).

Kullanım:
  python rapor_olustur.py girdi.json
  python rapor_olustur.py girdi.json --sablon "SEDDK Yakın İzleme Formatı_01092026.xlsx"
  python rapor_olustur.py girdi.json --ustyazi                     # xlsx + docx birlikte
  python rapor_olustur.py girdi.json --validate-only               # sadece doğrula
  python rapor_olustur.py --tara                                   # veri klasörlerini envanterler
  python rapor_olustur.py --ornek-girdi                             # boş girdi iskeleti yazar

girdi.json şeması (yıl/ay zorunlu; listeler opsiyonel — boşsa sayfaya satır yazılmaz):
{
  "yil": 2026, "ay": 10,
  "sirket_kodu": 1234,                    # yoksa assets/sirket_bilgileri.json'dan; o da yoksa UYARI
  "odemeler": [                           # '01-100+' sayfası — J sütunu 100.000,00 TL'yi AŞMALI
    {"alici": "...", "aciklama": "...", "dayanak": "...", "hesap_kodu": "...", "hesap_adi": "...",
     "tutar": 150000.00, "tarih": "05.10.2026", "odeme_ayi": 10, "dayanak_ek": "2026-10-05_fatura.pdf"}],
  "sozlesmeler": [                        # '02-Dış Hizmet' — 2026'da giderleştirilmiş TÜM sözleşmeler
    {"hizmet_konusu": "...", "ozet_icerik": "...", "amac": "...",
     "hesaplama_sekli": "Belirli Bedel", "birim_maliyeti": null,
     "saglayici": "...", "sozlesme_pdf": "sozlesme.pdf",
     "baslangic": "01.01.2026", "bitis": "31.12.2026", "sonlanma": null,
     "toplam_bedel": 600000.00, "tahakkuk": 50000.00, "odenen": 25000.00,
     "hesap_kodu": "...", "yk_karari": "Evet", "yk_pdf": "yk_karar.pdf"}],
  "ykk": [                                # '03- YKK' — ilgili ay içinde alınan kararlar
    {"karar_tarihi": "12.10.2026", "karar_sayisi": 1234, "gundem": "...", "karar": "...",
     "yk_pdf": "yk_karar.pdf"}],
  "notlar": ["..."],                      # üst yazıya ek beyan cümleleri (opsiyonel)
  "ustyazi": {"ref": "2026/123", "tarih": "15.11.2026"}   # opsiyonel; yoksa yer tutucu
}

Kurallar (Kurum hücre notlarından):
  * Yıl/Ay = GÖNDERİM DÖNEMİ. Ödeme satırında ayrıca K sütununa ödemenin ait olduğu ay.
  * Hesap kodu 'Sayısal' tipinde istenmiş; kod noktalı MSUG koduysa metin olarak yazılır (uyarı verir).
  * Sözleşme sürerken N sütunu (Sonlanma Tarihi) BOŞ bırakılır.
  * Dayanak belge / sözleşme PDF / YK kararı PDF ek olarak iletilecek → dosya klasörde aranır,
    bulunamazsa UYARI (üst yazı ek listesinde yine sayılır; teslim öncesi tamamlanmalı).
"""
import argparse
import glob
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from yazi_olustur import build as yazi_build, validate_input, collect_placeholders  # noqa: E402

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Side
except ImportError:
    sys.exit("openpyxl gerekli: pip install openpyxl")

AY_AD = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
         "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
TAR_RE = "%d.%m.%Y"
SHEET1, SHEET2, SHEET3 = "01-100+", "02-Dış Hizmet", "03- YKK"
ESIK = 100000.00  # 100 bin TL — 'üzeri' = aşan

# girdi anahtarı -> xlsx sütun sırası (A'dan itibaren)
S1_KEYS = ["yil", "ay", "sirket_adi", "sirket_kodu", "alici", "aciklama", "dayanak",
           "hesap_kodu", "hesap_adi", "tutar", "tarih", "odeme_ayi", "dayanak_ek"]
S2_KEYS = ["yil", "ay", "sirket_adi", "sirket_kodu", "hizmet_konusu", "ozet_icerik", "amac",
           "hesaplama_sekli", "birim_maliyeti", "saglayici", "sozlesme_pdf", "baslangic",
           "bitis", "sonlanma", "toplam_bedel", "tahakkuk", "odenen", "hesap_kodu",
           "yk_karari", "yk_pdf"]
S3_KEYS = ["yil", "ay", "sirket_adi", "sirket_kodu", "karar_tarihi", "karar_sayisi",
           "gundem", "karar", "yk_pdf"]

# hangi girdi anahtarı hangi klasörde aranan ek dosyası
EK_KLASOR = {"dayanak_ek": "ODEMELER", "sozlesme_pdf": "SOZLESMELER",
             "yk_pdf": "YKK",  # önce YKK, sonra SOZLESMELER'de de aranır
             }


def parse_tarih(s):
    """'GG.AA.YYYY' -> datetime; değilse None."""
    if s is None or s == "":
        return None
    if isinstance(s, datetime):
        return s
    try:
        return datetime.strptime(str(s).strip(), TAR_RE)
    except ValueError:
        return None


def fmt_tl(v):
    """1250000.0 -> '1.250.000,00'"""
    s = f"{v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def find_file(name, base):
    """base altındaki veri klasörlerinde dosyayı arar; bulunursa tam yol."""
    if not name:
        return None
    name = os.path.basename(str(name).strip())
    folders = [EK_KLASOR["dayanak_ek"], EK_KLASOR["sozlesme_pdf"], EK_KLASOR["yk_pdf"]]
    for fold in folders:
        d = os.path.join(base, fold)
        if not os.path.isdir(d):
            continue
        cand = os.path.join(d, name)
        if os.path.isfile(cand):
            return cand
        low = name.lower()
        for f in os.listdir(d):
            if f.lower() == low:
                return os.path.join(d, f)
    return None


# ---------- doğrulama ----------
def validate(girdi, sirket, base):
    errors, warnings = [], []

    yil, ay = girdi.get("yil"), girdi.get("ay")
    if not isinstance(yil, int) or yil < 2000:
        errors.append("yil eksik/hatalı (4 haneli sayı olmalı)")
    if not isinstance(ay, int) or not 1 <= ay <= 12:
        errors.append("ay eksik/hatalı (1-12)")

    sk = girdi.get("sirket_kodu", sirket.get("sirket_kodu"))
    if sk in (None, ""):
        warnings.append("sirket_kodu bilinmiyor — Veri Deseni D sütunu boş kalacak; "
                        "SEDDK şirket kodunu sorup tamamla")
    girdi.setdefault("_sk", sk)

    # --- ödemeler ---
    for i, o in enumerate(girdi.get("odemeler") or []):
        tag = f"odemeler[{i+1}]"
        for k in ("alici", "aciklama", "dayanak", "hesap_kodu", "hesap_adi", "tutar", "tarih"):
            if o.get(k) in (None, ""):
                errors.append(f"{tag}.{k} boş olamaz")
        t = o.get("tutar")
        if isinstance(t, (int, float)):
            if t <= ESIK:
                errors.append(f"{tag}.tutar={t} — 100 bin TL'yi AŞMALI; eşik ve altı ödeme "
                              "rapora alınmaz (ana faaliyet kalemleri zaten kapsam dışı)")
            if t < 0:
                errors.append(f"{tag}.tutar negatif")
        elif t not in (None, ""):
            errors.append(f"{tag}.tutar sayı olmalı")
        if parse_tarih(o.get("tarih")) is None and o.get("tarih") not in (None, ""):
            errors.append(f"{tag}.tarih GG.AA.YYYY formatında değil: {o.get('tarih')}")
        oa = o.get("odeme_ayi")
        if oa is None and parse_tarih(o.get("tarih")):
            o["odeme_ayi"] = parse_tarih(o["tarih"]).month
        elif oa is not None and not (isinstance(oa, int) and 1 <= oa <= 12):
            errors.append(f"{tag}.odeme_ayi 1-12 olmalı")
        hk = o.get("hesap_kodu")
        if isinstance(hk, str) and ("." in hk or hk.strip() == ""):
            warnings.append(f"{tag}.hesap_kodu='{hk}' metin olarak yazılacak (Veri Deseni "
                            "'Sayısal' istiyor; noktalı MSUG kodu varsa Kurumla biçim teyidi öner)")
        if o.get("dayanak_ek"):
            p = find_file(o["dayanak_ek"], base)
            if not p:
                warnings.append(f"{tag}.dayanak_ek='{o['dayanak_ek']}' veri klasörlerinde bulunamadı")

    # --- sözleşmeler ---
    for i, s in enumerate(girdi.get("sozlesmeler") or []):
        tag = f"sozlesmeler[{i+1}]"
        for k in ("hizmet_konusu", "ozet_icerik", "amac", "hesaplama_sekli", "saglayici",
                  "baslangic", "bitis", "toplam_bedel", "hesap_kodu", "yk_karari"):
            if s.get(k) in (None, ""):
                errors.append(f"{tag}.{k} boş olamaz")
        for k in ("baslangic", "bitis", "sonlanma"):
            if s.get(k) not in (None, "") and parse_tarih(s[k]) is None:
                errors.append(f"{tag}.{k} GG.AA.YYYY formatında değil: {s[k]}")
        tb = s.get("toplam_bedel")
        if tb not in (None, "") and not isinstance(tb, (int, float)):
            errors.append(f"{tag}.toplam_bedel sayı olmalı")
        for k in ("tahakkuk", "odenen"):
            if s.get(k) in (None, ""):
                s[k] = 0  # dönemde tahakkuk/ödeme yoksa 0 meşru
        if isinstance(tb, (int, float)) and tb > 500000:
            yk = str(s.get("yk_karari", "")).lower()
            if not yk.startswith("evet"):
                warnings.append(f"{tag}: toplam bedel 500.000 TL üzerinde ama yk_karari "
                                f"'{s.get('yk_karari')}' — 1937 sayılı Karar md.2 gereği Kurum "
                                "onayına tabi olmalı; onay yazısı (tür 4) kontrol edilmeli")
        hs = str(s.get("hesaplama_sekli", ""))
        if hs and not any(x in hs for x in ("Bağlı", "Belirli Bedel")):
            warnings.append(f"{tag}.hesaplama_sekli='{hs}' — Veri Deseni tip satırı "
                            "'Üretime Bağlı/Hasara Bağlı/…'a bağlı/Belirli Bedel' bekliyor")
        if str(s.get("yk_karari", "")).lower().startswith("evet") and not s.get("yk_pdf"):
            warnings.append(f"{tag}: yk_karari='Evet' ama yk_pdf verilmedi — T sütunu boş kalacak")
        if s.get("sonlanma") not in (None, "") and parse_tarih(s.get("bitis")) is None:
            errors.append(f"{tag}: bitiş tarihi yokken sonlanma tarihi verilmiş")
        for k in ("sozlesme_pdf", "yk_pdf"):
            if s.get(k):
                p = find_file(s[k], base)
                if not p:
                    warnings.append(f"{tag}.{k}='{s[k]}' veri klasörlerinde bulunamadı")

    # --- YKK ---
    for i, y in enumerate(girdi.get("ykk") or []):
        tag = f"ykk[{i+1}]"
        for k in ("karar_tarihi", "karar_sayisi", "gundem", "karar"):
            if y.get(k) in (None, ""):
                errors.append(f"{tag}.{k} boş olamaz")
        if y.get("karar_tarihi") and parse_tarih(y["karar_tarihi"]) is None:
            errors.append(f"{tag}.karar_tarihi GG.AA.YYYY değil: {y['karar_tarihi']}")
        ks = y.get("karar_sayisi")
        if ks not in (None, "") and not isinstance(ks, (int, float)):
            try:
                y["karar_sayisi"] = int(str(ks).strip())
            except ValueError:
                errors.append(f"{tag}.karar_sayisi sayı olmalı")
        kt = parse_tarih(y.get("karar_tarihi"))
        if kt and isinstance(yil, int):
            # sayfa başlığı: 'İlgili ay içerisinde alınan' — karar ayı gönderim dönemine ait olmalı
            if kt.year != yil or kt.month != ay:
                warnings.append(f"{tag}.karar_tarihi ({y['karar_tarihi']}) gönderim dönemi "
                                f"({ay}/{yil}) dışında — '03- YKK' sayfası ilgili ay kararlarını ister")
        if y.get("yk_pdf"):
            p = find_file(y["yk_pdf"], base)
            if not p:
                warnings.append(f"{tag}.yk_pdf='{y['yk_pdf']}' veri klasörlerinde bulunamadı")

    # ay 15 son gün kontrolü değil; teslim tarihi uyarısı üst yazıda:
    uy = girdi.get("ustyazi") or {}
    yt = parse_tarih(uy.get("tarih"))
    if isinstance(ay, int) and isinstance(yil, int) and yt:
        from datetime import date
        son = date(yil, ay, 1)
        while True:
            try:
                son = date(son.year, son.month, son.day + 1)
            except ValueError:
                break
        from datetime import timedelta
        kural = son + timedelta(days=15)  # biten ayın ertesi ayının 15'i
        if yt.date() > kural:
            warnings.append(f"Üst yazı tarihi ({uy['tarih']}) rapor son teslim gününü "
                            f"({tarih_str(kural)}) aşıyor")
    return errors, warnings


def tarih_str(d):
    return f"{d.day:02d}.{d.month:02d}.{d.year}"


# ---------- xlsx üretimi ----------
def put(ws, row, col, val, tip, thin):
    """Hücreye değer + biçim. tip: 'text'|'num'|'tutar'|'tarih'|'ay'"""
    c = ws.cell(row=row, column=col)
    c.border = thin
    if val is None or val == "":
        return
    if tip == "num":
        try:
            c.value = int(val) if float(val) == int(float(val)) else float(val)
        except (TypeError, ValueError):
            c.value = str(val)  # noktalı hesap kodu vb.
        c.number_format = "0" if isinstance(c.value, int) else "0.00"
        c.alignment = Alignment(horizontal="right")
    elif tip == "tutar":
        c.value = float(val)
        c.number_format = "#,##0.00"
        c.alignment = Alignment(horizontal="right")
    elif tip == "tarih":
        d = parse_tarih(val)
        if d:
            c.value = d
            c.number_format = "DD.MM.YYYY"
        else:
            c.value = str(val)
        c.alignment = Alignment(horizontal="center")
    else:
        c.value = str(val)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def fill_sheet(ws, keys, rows, insert_at, types, thin, sirket_adi, yil, ay, sk):
    """Tip satırının altına veri satırlarını ekler; şirket/yıl/ay ön sütunlarını otomatik doldurur."""
    n = len(rows)
    if n:
        ws.insert_rows(insert_at, n)
    for i, r in enumerate(rows):
        for j, k in enumerate(keys):
            col = j + 1
            if k == "sirket_adi":
                val = sirket_adi
            elif k == "yil":
                val = yil
            elif k == "ay":
                val = ay
            elif k == "sirket_kodu":
                val = sk
            else:
                val = r.get(k)
            put(ws, insert_at + i, col, val, types[j], thin)


def build_xlsx(girdi, sablon, out_path, sirket):
    wb = load_workbook(sablon)
    thin = Border(*[Side(style="thin")] * 4)

    sirket_adi = sirket.get("sirket_adi", "")
    yil, ay = girdi["yil"], girdi["ay"]
    sk = girdi.get("_sk")

    # tip dizileri: A..M (13) / A..T (20) / A..I (9)
    t1 = ["num", "num", "text", "num", "text", "text", "text", "num", "text",
          "tutar", "tarih", "num", "text"]
    t2 = ["num", "num", "text", "num", "text", "text", "text", "text", "text", "text",
          "text", "tarih", "tarih", "tarih", "tutar", "tutar", "tutar", "num", "text", "text"]
    t3 = ["num", "num", "text", "num", "tarih", "num", "text", "text", "text"]

    fill_sheet(wb[SHEET1], S1_KEYS, girdi.get("odemeler") or [], 3, t1, thin, sirket_adi, yil, ay, sk)
    fill_sheet(wb[SHEET2], S2_KEYS, girdi.get("sozlesmeler") or [], 4, t2, thin, sirket_adi, yil, ay, sk)
    fill_sheet(wb[SHEET3], S3_KEYS, girdi.get("ykk") or [], 5, t3, thin, sirket_adi, yil, ay, sk)
    wb.save(out_path)

    # --- üretim sonrası kendi kendini doğrula ---
    errs = []
    wb2 = load_workbook(out_path)
    ws1 = wb2[SHEET1]
    r = 3
    while ws1.cell(row=r, column=1).value not in (None, ""):
        tutar = ws1.cell(row=r, column=10).value
        if not isinstance(tutar, (int, float)) or tutar <= ESIK:
            errs.append(f"{SHEET1} satır {r}: J={tutar} — 100 bin TL üzeri olmalı")
        r += 1
    n1 = r - 3
    beklenen1 = len(girdi.get("odemeler") or [])
    if n1 != beklenen1:
        errs.append(f"{SHEET1}: beklenen {beklenen1} satır, yazılan {n1}")
    ws2 = wb2[SHEET2]
    r = 4
    while ws2.cell(row=r, column=1).value not in (None, ""):
        r += 1
    if (r - 4) != len(girdi.get("sozlesmeler") or []):
        errs.append(f"{SHEET2}: satır sayısı uyuşmaz")
    ws3 = wb2[SHEET3]
    r = 5
    while ws3.cell(row=r, column=1).value not in (None, ""):
        r += 1
    if (r - 5) != len(girdi.get("ykk") or []):
        errs.append(f"{SHEET3}: satır sayısı uyuşmaz")
    return errs, n1


# ---------- üst yazı ----------
def build_ustyazi_girdi(girdi, sirket, xlsx_name):
    yil, ay = girdi["yil"], girdi["ay"]
    ayad = AY_AD[ay]
    kunye = sirket.get("raporlama_kunyesi", {})

    ilgi = [
        sirket["karar_kunyesi"]["ana_ilgi"],
        kunye.get("ilgi", "Kurumunuzun E-51048673-010.05-5045010 sayılı yazısı (Kuruma Yapılacak Raporlamalar Hk.)."),
    ]

    odemeler = girdi.get("odemeler") or []
    sozlesmeler = girdi.get("sozlesmeler") or []
    ykk = girdi.get("ykk") or []
    toplam = sum(float(o["tutar"]) for o in odemeler if isinstance(o.get("tutar"), (int, float)))

    govde = [
        f"İlgi (a) yazınız ile Şirketimizin, Sigortacılık ve Özel Emeklilik Düzenleme ve Denetleme "
        f"Kurulu'nun 23.09.2026 tarih ve 1937 sayılı Kararı uyarınca yakın izleme kapsamına alındığı "
        f"bildirilmiştir.",
        f"İlgi (b) yazınız ile anılan Karar uyarınca; ana sigortacılık faaliyeti kapsamındaki ödemeler "
        f"hariç olmak üzere 100 bin TL üzeri ödemelere, yönetim kurulu kararlarına ve dış hizmet "
        f"alımlarına ilişkin bilgilerin, işbu yazı ekindeki şekil ve formatta, Şirketimiz iç denetim "
        f"müdürü ve iç sistemlerden sorumlu yönetim kurulu üyesi tarafından imzalı yazı ile biten aya "
        f"ilişkin raporun her ayın 15'ine kadar düzenli olarak Kurumunuza sunulması hususunda bilgi "
        f"verilmiştir.",
        f"Bu çerçevede {ayad} {yil} dönemine ilişkin rapor ekte sunulmuştur.",
    ]
    if odemeler:
        govde.append(
            f"Söz konusu dönemde ana sigortacılık faaliyeti kapsamındaki ödemeler hariç, 100 bin TL "
            f"üzeri {len(odemeler)} adet ödeme gerçekleşmiş olup ödemelere ilişkin bilgiler Veri "
            f"Deseni'nin '{SHEET1}' sayfasında, dayanak belgeleri ise ekte sunulmuştur. "
            f"Dönem ödeme toplamı {fmt_tl(toplam)} TL'dir.")
    else:
        govde.append(
            f"Söz konusu dönemde, ana sigortacılık faaliyeti kapsamındaki ödemeler hariç, 100 bin TL "
            f"üzeri ödeme gerçekleşmemiştir.")
    if sozlesmeler:
        govde.append(
            f"Dışarıdan hizmet alımlarına ilişkin olarak, 2026 yılı mali tablolarında giderleştirilmiş "
            f"sözleşmelere ait bilgiler Veri Deseni'nin '{SHEET2}' sayfasında sunulmuş olup sözleşme "
            f"ve yönetim kurulu kararına ilişkin belgeler ekte yer almaktadır.")
    else:
        govde.append(
            f"2026 yılı mali tablolarında giderleştirilmiş, rapora konu dış hizmet alımı sözleşmesi "
            f"bulunmamaktadır.")
    if ykk:
        govde.append(
            f"Dönem içinde alınan {len(ykk)} adet yönetim kurulu kararına ilişkin bilgiler Veri "
            f"Deseni'nin '{SHEET3}' sayfasında sunulmuş, kararların PDF suretleri ekte yer almaktadır.")
    else:
        govde.append("Dönem içinde yönetim kurulu kararı alınmamıştır.")
    govde.extend(girdi.get("notlar") or [])
    govde.append("Bilgilerinize arz ederiz.")

    ekler = [f"Veri Deseni ({xlsx_name})"]
    if odemeler:
        ekler.append(f"Ödemelere ilişkin dayanak belgeler ({len(odemeler)} adet)")
    if sozlesmeler:
        ekler.append(f"Dış hizmet alımı sözleşmeleri ve varsa yönetim kurulu kararları "
                     f"({len(sozlesmeler)} adet)")
    if ykk:
        ekler.append(f"Yönetim kurulu kararlarının PDF suretleri ({len(ykk)} adet)")

    uy = girdi.get("ustyazi") or {}
    data = {
        "ref": uy.get("ref", "[●YYYY/NNN]"),
        "tarih": uy.get("tarih", "[●GG.AA.YYYY]"),
        "konu": f"Yakın İzleme Kapsamında {ayad} {yil} Dönemine İlişkin Rapor Sunumu",
        "ilgi": ilgi,
        "govde": govde,
        "ekler": ekler,
        "imza": kunye.get("imzacilar") or sirket.get("varsayilan_imzacilar"),
    }
    return data


# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="SEDDK yakın izleme aylık raporu üretici (v4)")
    ap.add_argument("girdi", nargs="?", help="Rapor girdisi JSON")
    ap.add_argument("--sablon", default=None, help="Kurum Veri Deseni şablonu (.xlsx)")
    ap.add_argument("--cikti", default=None, help="Çıktı xlsx yolu")
    ap.add_argument("--ustyazi", nargs="?", const="auto", default=None,
                    help="Üst yazı .docx de üret (yol verilmezse otomatik ad)")
    ap.add_argument("--sirket", default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                      "..", "assets", "sirket_bilgileri.json"))
    ap.add_argument("--veri-dizini", default=None,
                    help="Veri klasörlerinin kökü (varsayılan: girdinin bulunduğu dizin)")
    ap.add_argument("--validate-only", action="store_true")
    ap.add_argument("--tara", action="store_true", help="Veri klasörlerini envanterle")
    ap.add_argument("--ornek-girdi", action="store_true", help="Boş girdi iskeleti yaz (ornek.json)")
    a = ap.parse_args()

    if a.tara:
        base = a.veri_dizini or os.getcwd()
        for fold in (EK_KLASOR["dayanak_ek"], EK_KLASOR["sozlesme_pdf"], EK_KLASOR["yk_pdf"]):
            d = os.path.join(base, fold)
            print(f"--- {fold} ---")
            if not os.path.isdir(d):
                print("  (klasör yok)")
                continue
            files = sorted(os.listdir(d))
            if not files:
                print("  (boş)")
            for f in files:
                print("  ", f)
        tpl = glob.glob(os.path.join(base, "SEDDK*Format*.xlsx"))
        print("Şablon:", tpl[0] if tpl else "(bulunamadı — --sablon verin)")
        return

    if a.ornek_girdi:
        isk = {
            "yil": 2026, "ay": 10, "sirket_kodu": None,
            "odemeler": [], "sozlesmeler": [], "ykk": [],
            "notlar": [], "ustyazi": {"ref": "[●YYYY/NNN]", "tarih": "[●GG.AA.YYYY]"},
        }
        out = a.girdi or "ornek_girdi.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump(isk, f, ensure_ascii=False, indent=2)
        print("İskelet yazıldı:", out)
        return

    if not a.girdi:
        ap.error("girdi.json gerekli (veya --tara / --ornek-girdi)")
    if not os.path.isfile(a.girdi):
        sys.exit(f"Girdi bulunamadı: {a.girdi}")
    with open(a.girdi, encoding="utf-8") as f:
        girdi = json.load(f)
    with open(a.sirket, encoding="utf-8") as f:
        sirket = json.load(f)

    base = a.veri_dizini or os.path.dirname(os.path.abspath(a.girdi))

    sablon = a.sablon
    if not sablon:
        cand = glob.glob(os.path.join(base, "SEDDK*Format*.xlsx"))
        if cand:
            sablon = cand[0]
        else:
            sys.exit("Şablon bulunamadı: --sablon ile 'SEDDK Yakın İzleme Formatı' xlsx'ini verin")

    errors, warnings = validate(girdi, sirket, base)
    for w in warnings:
        print(f"Uyarı: {w}", file=sys.stderr)
    if errors:
        for e in errors:
            print(f"Hata: {e}", file=sys.stderr)
        sys.exit(1)
    if a.validate_only:
        print("Doğrulama başarılı.")
        return

    stem = f"SEDDK_YakinIzleme_VeriDeseni_{girdi['yil']}-{girdi['ay']:02d}"
    xlsx = a.cikti or os.path.join(base, stem + ".xlsx")
    post_errs, n1 = build_xlsx(girdi, sablon, xlsx, sirket)
    if post_errs:
        for e in post_errs:
            print(f"Hata (üretim sonrası): {e}", file=sys.stderr)
        sys.exit(1)
    print(f"Veri deseni üretildi: {xlsx}")
    print(f"  {SHEET1}: {len(girdi.get('odemeler') or [])} ödeme satırı")
    print(f"  {SHEET2}: {len(girdi.get('sozlesmeler') or [])} sözleşme satırı")
    print(f"  {SHEET3}: {len(girdi.get('ykk') or [])} YK kararı satırı")

    if a.ustyazi:
        docx = a.ustyazi if a.ustyazi != "auto" else os.path.join(
            base, f"SEDDK_YakinIzleme_Rapor_{girdi['yil']}-{girdi['ay']:02d}.docx")
        data = build_ustyazi_girdi(girdi, sirket, os.path.basename(xlsx))
        yerr, ywarn = validate_input(data, sirket)
        for w in ywarn:
            print(f"Üst yazı uyarısı: {w}", file=sys.stderr)
        if yerr:
            for e in yerr:
                print(f"Üst yazı hatası: {e}", file=sys.stderr)
            sys.exit(1)
        yazi_build(data, sirket, docx)
        print(f"Üst yazı üretildi: {docx}")
        left = collect_placeholders(data, sirket)
        if left:
            print("Doldurulacak yer tutucular:")
            for x in left:
                print("  ", x)
    print("Sonraki adım: python scripts/dogrula.py <üst yazı .docx>")


if __name__ == "__main__":
    main()