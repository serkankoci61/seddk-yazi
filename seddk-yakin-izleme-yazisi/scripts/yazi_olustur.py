#!/usr/bin/env python3
"""SEDDK'ya gönderilecek resmi yazıyı, şirketin mevcut yazı formatında .docx olarak üretir (v3).

Biçim (şirketin Pert/Sovtaj, İç Denetim Envanteri ve ORNEK_Sozlesme_Onay_Talebi örneklerinden):
Arial 12, A4, 2,5 cm kenar boşluğu; solda muhatap (3 satır, kalın), sağda Tarih ve Ref;
Konu / İlgi; iki yana yaslı, ilk satır girintili, ~1,15 satır aralıklı gövde; ortalı
"Saygılarımızla," + şirket unvanı + yan yana iki imza; "Ek:" listesi.

v3 değişiklikleri (v2'ye göre):
- Satır aralığı 1.15 (gerçek örnek yazıyla uyumlu; 3 sayfa yerine 2 sayfa)
- keep_with_next zinciri: kapanış → Saygı → şirket → imza → Ek listesi bölünmeye karşı bağlı;
  tablodan önceki paragraf tabloyla aynı sayfada kalır
- İmza bloğu tek tabloda (imza alanı + ad + unvan aynı hücrede; cantSplit)
- Yeni doğrulamalar: İlgi (x) atıf/ilgi listesi uyumu, Ek-N atıf/ek listesi uyumu,
  tablo toplam satırı aritmetik kontrolü, sayı biçimi (Türk formatı) kontrolü,
  kapanış cümlesi kontrolü, karar tarihi öncesi yazı tarihi uyarısı
- Tablo başlık dolgusu E7E6E6 (gerçek örnek yazıyla bire bir uyum)

Kullanım:
    python yazi_olustur.py girdi.json cikti.docx [--sirket sirket_bilgileri.json] [--validate-only]
    python scripts/dogrula.py cikti.docx        # üretim sonrası sayfa yerleşimi kontrolü (pymupdf)

girdi.json (* zorunlu):
{
  "ref":    "2026/123",                 # şirket Ref numarası (yoksa [●] yer tutucusu)
  "tarih":  "30.09.2026",               # GG.AA.YYYY
  "konu":   "...",                      # *
  "ilgi":   ["..."],                    # * en az 1; birden fazlaysa a), b) otomatik harflenir
                                        #   TEK ilgi harflenmez (resmî yazışma kuralı) ve
                                        #   gövdede atıf "İlgi yazınız ile" olmalıdır
  "dikkat": "...",                      # opsiyonel, muhatabın altında italik satır
  "govde":  [ ... ],                    # *
  "ekler":  ["..."],                    # opsiyonel; gövdede "Ek-1"/"ekte" atfıyla eşleşmeli
  "imza":   [{"ad": "...", "unvan": "..."}]   # verilmezse şirket dosyasındaki varsayılan
}

govde öğeleri:
  "düz metin"                           -> iki yana yaslı paragraf (ilk satır girintili);
                                           "... arz ederiz." ile biten kısa cümle girintisiz kapanış olur
  {"baslik": "1. ..."}                  -> kalın alt başlık
  {"liste": [...], "tip": "harf"|"rakam"|"madde"}
  {"tablo": {"basliklar": [...], "satirlar": [[...]], "genislikler_cm": [..],
             "sayi_sutunlari": [idx], "toplam_satiri": [...]}}
  {"not": "küçük punto açıklama"}

sayi_sutunlari hücreleri "1.250.000,00" Türk formatında olmalıdır; toplam_satiri varsa
script satırların toplamıyla aritmetik olarak karşılaştırır (resmî yazıda yanlış toplam
Kurum nezdinde güven kaybıdır; uyuşmazlıkta üretim DURUR).

[●...] biçimindeki yer tutucular sarı vurgulanır; gönderimden önce doldurulmalıdır.
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "Arial"
PLACEHOLDER = re.compile(r"(\[●[^\]]*\])")
DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
REF_RE = re.compile(r"^(\d{4}/\d+|\[●[^\]]*\])$")
KAPANIS = re.compile(r"arz ederiz\.$")
TL_RE = re.compile(r"^\d{1,3}(\.\d{3})*,\d{2}$")
ILGI_ATIF_RE = re.compile(r"İlgi \(([a-z])\)")
EK_NO_RE = re.compile(r"Ek[- ]?(\d+)")
EK_ATIF_RE = re.compile(r"\b(?:[Ee]kte\b|[Ee]k[- ]?\d|[Ee]k'|\bilişik|Ekler\b)")


def parse_tl(s):
    """"1.250.000,00" -> 1250000.00; format dışıysa None."""
    s = str(s).strip().replace(" TL", "")
    if TL_RE.match(s):
        return float(s.replace(".", "").replace(",", "."))
    return None


# ---------- yardımcılar ----------
def set_font(run, size=None, bold=None, italic=None, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(a), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_text(par, text, size=12, bold=False, italic=False, color=None):
    """Metni ekler; [●...] yer tutucularını sarı vurgular."""
    for part in PLACEHOLDER.split(text):
        if not part:
            continue
        r = par.add_run(part)
        set_font(r, size, bold, italic, color)
        if PLACEHOLDER.fullmatch(part):
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return par


def fmt(par, align=None, before=0, after=6, left=None, first=None, hanging=None,
        keep_next=False, line=1.15):
    pf = par.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if align is not None:
        par.alignment = align
    if left is not None:
        pf.left_indent = Cm(left)
    if first is not None:
        pf.first_line_indent = Cm(first)
    if hanging is not None:
        pf.left_indent = Cm(hanging)
        pf.first_line_indent = Cm(-hanging)
    pf.keep_with_next = keep_next
    return par


def shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def cell_margins(table, top=50, bottom=50, left=80, right=80):
    tblpr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblpr.append(mar)


def no_split_row(row):
    trpr = row._tr.get_or_add_trPr()
    e = OxmlElement("w:cantSplit")
    trpr.append(e)


def repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    trpr.append(e)


def no_borders(table):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for k in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{k}")
        e.set(qn("w:val"), "nil")
        borders.append(e)
    tblpr.append(borders)


def fixed_layout(table):
    tblpr = table._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblpr.append(lay)


def set_cell_vertical_align(cell, align="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), align)
    tcPr.append(vAlign)


# ---------- gövde öğeleri ----------
def render_liste(doc, item):
    tip = item.get("tip", "harf")
    for i, txt in enumerate(item["liste"]):
        if tip == "harf":
            mark = f"{chr(ord('a') + i)})"
        elif tip == "rakam":
            mark = f"{i + 1}."
        else:
            mark = "•"
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(2.0))
        add_text(p, f"{mark}\t{txt}")
        fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, after=3, left=2.0, first=-0.75)


def render_tablo(doc, t, usable_cm):
    heads = t["basliklar"]
    rows = t["satirlar"]
    toplam = t.get("toplam_satiri")
    n = len(heads)
    if n == 0:
        raise ValueError("Tablo başlıkları boş olamaz")
    for ri, r in enumerate(rows):
        if len(r) != n:
            raise ValueError(f"Tablo satır {ri+1} sütun sayısı ({len(r)}) başlıklarla ({n}) uyuşmuyor")
    if toplam and len(toplam) != n:
        raise ValueError(f"Toplam satırı sütun sayısı ({len(toplam)}) başlıklarla ({n}) uyuşmuyor")

    widths = t.get("genislikler_cm")
    if not widths:
        widths = [usable_cm / n] * n
    if len(widths) != n:
        raise ValueError(f"genislikler_cm uzunluğu ({len(widths)}) başlık sayısıyla ({n}) uyuşmuyor")
    scale = usable_cm / sum(widths)
    widths = [w * scale for w in widths]

    table = doc.add_table(rows=1, cols=n)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    fixed_layout(table)
    cell_margins(table, top=40, bottom=40, left=60, right=60)
    for i, col in enumerate(table.columns):
        col.width = Cm(widths[i])

    def fill(row, values, bold=False, fill_hex=None, align_right_cols=()):
        for i, v in enumerate(values):
            c = row.cells[i]
            c.width = Cm(widths[i])
            set_cell_vertical_align(c, "center")
            p = c.paragraphs[0]
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            add_text(p, str(v), size=9 if n > 5 else 10, bold=bold)
            align = WD_ALIGN_PARAGRAPH.RIGHT if i in align_right_cols else WD_ALIGN_PARAGRAPH.LEFT
            if i == 0 and not (i in align_right_cols):
                align = WD_ALIGN_PARAGRAPH.CENTER if str(v).isdigit() else WD_ALIGN_PARAGRAPH.LEFT
            fmt(p, align, after=0, before=0, line=1.05)
            if fill_hex:
                shade(c, fill_hex)

    money_cols = set(t.get("sayi_sutunlari", []))
    fill(table.rows[0], heads, bold=True, fill_hex="E7E6E6")  # gerçek örnek yazıyla uyumlu gri
    repeat_header(table.rows[0])
    no_split_row(table.rows[0])
    for r in rows:
        row = table.add_row()
        no_split_row(row)
        fill(row, r, align_right_cols=money_cols)
    if toplam:
        row = table.add_row()
        no_split_row(row)
        fill(row, toplam, bold=True, fill_hex="FFF2CC", align_right_cols=money_cols)
    sp = doc.add_paragraph()
    fmt(sp, after=6, line=1.0)


# ---------- doğrulama ----------
def govde_blob(data):
    """Gövde öğelerinin tamamını tek metne döker (atıf kontrolleri için)."""
    parts = []
    for item in data.get("govde", []):
        if isinstance(item, str):
            parts.append(item)
        elif isinstance(item, dict):
            if "liste" in item:
                parts.extend(str(x) for x in item["liste"])
            if "not" in item:
                parts.append(str(item["not"]))
            if "baslik" in item:
                parts.append(str(item["baslik"]))
            if "tablo" in item:
                t = item["tablo"]
                parts.extend(str(h) for h in t.get("basliklar", []))
                for row in t.get("satirlar", []):
                    parts.extend(str(c) for c in row)
                if t.get("toplam_satiri"):
                    parts.extend(str(c) for c in t["toplam_satiri"])
    return " ".join(parts)


def validate_input(data, sirket):
    """Girdi JSON'unu doğrular; (hata, uyarı) listeleri döner."""
    errors = []
    warnings = []

    for k in ("konu", "ilgi", "govde"):
        if not data.get(k):
            errors.append(f"Zorunlu alan eksik: {k}")

    karar_tarihi = sirket.get("karar_kunyesi", {}).get("karar_tarihi")

    if data.get("tarih"):
        if not DATE_RE.match(data["tarih"]):
            errors.append(f"Tarih formatı hatalı (GG.AA.YYYY olmalı): {data['tarih']}")
        else:
            try:
                yazi_t = datetime.strptime(data["tarih"], "%d.%m.%Y")
                if karar_tarihi:
                    try:
                        karar_t = datetime.strptime(karar_tarihi, "%d.%m.%Y")
                        if yazi_t < karar_t:
                            warnings.append(
                                f"Yazı tarihi ({data['tarih']}) karar tarihinden ({karar_tarihi}) önce olamaz")
                    except ValueError:
                        pass
            except ValueError:
                errors.append(f"Geçersiz tarih: {data['tarih']}")
    else:
        warnings.append("tarih alanı yok; [●GG.AA.YYYY] yer tutucusu kullanılacak")

    if data.get("ref") and not REF_RE.match(str(data["ref"])):
        warnings.append(f"Ref formatı beklenenden farklı (YYYY/NNN): {data['ref']}")

    if isinstance(data.get("ilgi"), list):
        if len(data["ilgi"]) == 0:
            errors.append("ilgi listesi boş olamaz")
        for i, txt in enumerate(data["ilgi"]):
            if not isinstance(txt, str) or not txt.strip():
                errors.append(f"ilgi[{i}] boş veya geçersiz")
    elif data.get("ilgi") is not None:
        errors.append("ilgi bir liste olmalıdır")

    if isinstance(data.get("govde"), list):
        if len(data["govde"]) == 0:
            errors.append("govde listesi boş olamaz")
        blob = govde_blob(data)

        # --- İlgi atıf uyumu ---
        ilgi = data.get("ilgi") or []
        atiflar = ILGI_ATIF_RE.findall(blob)
        for harf in atiflar:
            idx = ord(harf) - ord("a")
            if idx + 1 > len(ilgi):
                errors.append(
                    f"Gövdede 'İlgi ({harf})' atfı var ama ilgi listesinde yalnız {len(ilgi)} öğe var")
        if len(ilgi) == 1 and "İlgi (a)" in blob:
            warnings.append(
                "Tek ilgi var; script ilgiyi harflemez. Gövdedeki atfı 'İlgi yazınız ile ...' yapın "
                "(resmî yazışmada tek ilgi harflenmez)")
        if len(ilgi) > 1 and re.search(r"İlgi yazınız", blob):
            warnings.append(
                f"{len(ilgi)} ilgi var ama gövdede harfsiz 'İlgi yazınız' atfı geçiyor; 'İlgi (a) yazınız' kullanın")

        # --- Ek atıf uyumu ---
        ekler = data.get("ekler") or []
        ek_nolar = EK_NO_RE.findall(blob)
        if ek_nolar:
            mx = max(int(x) for x in ek_nolar)
            if not ekler:
                warnings.append("Gövdede 'Ek-...' atfı var ama ekler listesi boş")
            elif mx > len(ekler):
                errors.append(f"Gövdede Ek-{mx} atfı var ama ekler listesinde {len(ekler)} madde var")
        if ekler and not ek_nolar and not EK_ATIF_RE.search(blob):
            warnings.append("Ekler listesi dolu ama gövdede hiçbir ek atfı ('ekte sunulmuştur', 'Ek-1' vb.) yok")
        if not ekler and EK_ATIF_RE.search(blob):
            warnings.append("Gövdede ek atfı geçiyor ama ekler listesi yok")

        # --- Tablo aritmetik ve biçim kontrolleri ---
        for gi, item in enumerate(data["govde"]):
            if not isinstance(item, dict):
                continue
            if not set(item.keys()) & {"baslik", "liste", "tablo", "not"}:
                errors.append(f"govde[{gi}] tanınmayan anahtar: {set(item.keys())}")
            if "tablo" in item:
                t = item["tablo"]
                if not isinstance(t, dict) or "basliklar" not in t or "satirlar" not in t:
                    errors.append(f"govde[{gi}].tablo basliklar ve satirlar içermeli")
                    continue
                for idx, col in enumerate(t.get("sayi_sutunlari", [])):
                    for ri, row in enumerate(t["satirlar"]):
                        if col >= len(row):
                            continue
                        v = str(row[col])
                        if PLACEHOLDER.fullmatch(v.strip()):
                            continue
                        if not v.strip():
                            continue  # boş hücre sessiz geçilir
                        if not TL_RE.match(v.strip()):
                            warnings.append(
                                f"govde[{gi}] tablo satır {ri+1}, sütun {col+1}: '{v}' Türk para "
                                f"formatında değil (örn. 1.250.000,00)")
                toplam = t.get("toplam_satiri")
                if toplam:
                    for col in t.get("sayi_sutunlari", []):
                        vals = []
                        ok = True
                        for row in t["satirlar"]:
                            if col >= len(row):
                                ok = False
                                break
                            pv = parse_tl(row[col])
                            if pv is None:
                                ok = False
                                break
                            vals.append(pv)
                        if not ok or col >= len(toplam):
                            continue
                        tv = parse_tl(toplam[col])
                        if tv is not None and abs(sum(vals) - tv) > 0.005:
                            errors.append(
                                f"govde[{gi}] tablo sütun {col+1}: satır toplamı "
                                f"{sum(vals):,.2f} ama toplam satırı {tv:,.2f} — DÜZELT")

        # --- Kapanış cümlesi ---
        son = data["govde"][-1] if data["govde"] else None
        if isinstance(son, str) and not KAPANIS.search(son):
            warnings.append(
                "Son gövde öğesi '... arz ederiz.' kapanışı ile bitmiyor — resmî yazıda kapanış zorunludur")
        elif son is not None and not isinstance(son, str):
            warnings.append("Son gövde öğesi metin değil (tablo/liste); kapanış cümlesi eksik olabilir")
    elif data.get("govde") is not None:
        errors.append("govde bir liste olmalıdır")

    if data.get("imza"):
        for i, sg in enumerate(data["imza"]):
            if not isinstance(sg, dict) or "ad" not in sg or "unvan" not in sg:
                errors.append(f"imza[{i}] ad ve unvan içermeli")

    if not sirket.get("muhatap_satirlari"):
        errors.append("şirket dosyasında muhatap_satirlari eksik")
    if not sirket.get("varsayilan_imzacilar"):
        warnings.append("varsayilan_imzacilar eksik; imza bloğu boş kalabilir")

    return errors, warnings


def collect_placeholders(data, sirket):
    blob = json.dumps(data, ensure_ascii=False) + json.dumps(sirket, ensure_ascii=False)
    return sorted(set(PLACEHOLDER.findall(blob)))


# ---------- ana üretim ----------
def build(data, sirket, out_path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.top_margin = sec.bottom_margin = Cm(2.5)
    sec.left_margin = sec.right_margin = Cm(2.5)
    usable = 16.0

    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(12)

    # Antetli kağıt payı
    sp = doc.add_paragraph()
    fmt(sp, after=0, line=1.0)
    sp.paragraph_format.space_before = Pt(sirket.get("ust_bosluk_pt", 24))

    # Muhatap (sol) + Tarih/Ref (sağ)
    hdr = doc.add_table(rows=1, cols=2)
    no_borders(hdr)
    hdr.autofit = False
    fixed_layout(hdr)
    widths = (11.0, 5.0)
    for i, c in enumerate(hdr.rows[0].cells):
        c.width = Cm(widths[i])
    hdr.columns[0].width, hdr.columns[1].width = Cm(widths[0]), Cm(widths[1])
    left, right = hdr.rows[0].cells
    first = True
    for line in sirket["muhatap_satirlari"]:
        p = left.paragraphs[0] if first else left.add_paragraph()
        first = False
        add_text(p, line, bold=True)
        fmt(p, WD_ALIGN_PARAGRAPH.LEFT, after=0, line=1.15)
    p = right.paragraphs[0]
    add_text(p, "Tarih: " + data.get("tarih", "[●GG.AA.YYYY]"), bold=True)
    fmt(p, WD_ALIGN_PARAGRAPH.LEFT, after=0, line=1.15)
    p = right.add_paragraph()
    add_text(p, "Ref: " + data.get("ref", "[●YYYY/NNN]"), bold=True)
    fmt(p, WD_ALIGN_PARAGRAPH.LEFT, after=0, line=1.15)

    if data.get("dikkat"):
        p = doc.add_paragraph()
        add_text(p, f"Dikkat: {data['dikkat']}", size=11, italic=True)
        fmt(p, WD_ALIGN_PARAGRAPH.LEFT, before=6, after=0, line=1.15)

    sp = doc.add_paragraph()
    fmt(sp, after=0, line=1.0)

    # Konu
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.5))
    add_text(p, "Konu\t", bold=True)
    add_text(p, ": " + data["konu"])
    fmt(p, WD_ALIGN_PARAGRAPH.LEFT, after=2, hanging=1.5, line=1.15)

    # İlgi
    ilgi = data["ilgi"]
    ind = 1.5 + (0.5 if len(ilgi) > 1 else 0)
    for i, txt in enumerate(ilgi):
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(1.5))
        if i == 0:
            add_text(p, "İlgi\t", bold=True)
            add_text(p, ": ")
        else:
            add_text(p, "\t  ")
        if len(ilgi) > 1:
            add_text(p, f"{chr(ord('a') + i)}) ")
        add_text(p, txt)
        fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, after=2, left=ind, first=-ind, line=1.15)

    sp = doc.add_paragraph()
    fmt(sp, after=0, line=1.0)

    # Gövde
    govde = data["govde"]
    for gi, item in enumerate(govde):
        sonraki_tablo = (gi + 1 < len(govde) and isinstance(govde[gi + 1], dict)
                         and "tablo" in govde[gi + 1])
        if isinstance(item, str):
            p = doc.add_paragraph()
            add_text(p, item)
            if KAPANIS.search(item) and len(item) < 80:
                fmt(p, WD_ALIGN_PARAGRAPH.LEFT, before=8, after=6, keep_next=True)
            else:
                # tabloyu izleyen giriş cümlesi tabloyla aynı sayfada kalsın
                fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, after=4, first=1.25, keep_next=sonraki_tablo)
        elif "baslik" in item:
            p = doc.add_paragraph()
            add_text(p, item["baslik"], bold=True)
            fmt(p, WD_ALIGN_PARAGRAPH.LEFT, before=8, after=4, keep_next=True)
        elif "liste" in item:
            render_liste(doc, item)
            sp = doc.add_paragraph()
            fmt(sp, after=2, line=1.0)
        elif "tablo" in item:
            render_tablo(doc, item["tablo"], usable)
        elif "not" in item:
            p = doc.add_paragraph()
            add_text(p, item["not"], size=10, italic=True)
            fmt(p, WD_ALIGN_PARAGRAPH.JUSTIFY, after=8, line=1.15)

    # Saygılarımızla + şirket unvanı + imzalar + Ek — keep_with_next zinciriyle tek blok
    p = doc.add_paragraph()
    add_text(p, "Saygılarımızla,")
    fmt(p, WD_ALIGN_PARAGRAPH.CENTER, before=12, after=0, keep_next=True, line=1.15)
    p = doc.add_paragraph()
    add_text(p, sirket.get("sirket_adi_kisa", sirket.get("sirket_adi", "")), bold=True)
    fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=0, keep_next=True, line=1.15)

    imzalar = data.get("imza") or sirket.get("varsayilan_imzacilar", [])
    if imzalar:
        # Tek tablo: imza alanı (boşluk) + ad + unvan aynı hücrede; cantSplit bölünmeyi engeller
        sig = doc.add_table(rows=1, cols=len(imzalar))
        sig.alignment = WD_TABLE_ALIGNMENT.CENTER
        sig.autofit = False
        no_borders(sig)
        fixed_layout(sig)
        no_split_row(sig.rows[0])
        w = usable / len(imzalar)
        for i, col in enumerate(sig.columns):
            col.width = Cm(w)
        for i, sg in enumerate(imzalar):
            c = sig.rows[0].cells[i]
            c.width = Cm(w)
            p = c.paragraphs[0]
            add_text(p, "")
            fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=0, before=0, line=1.0)
            p.paragraph_format.space_before = Pt(36)  # ıslak/e-imza için dikey boşluk
            p = c.add_paragraph()
            add_text(p, sg["ad"], bold=True)
            fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=0, line=1.15)
            p = c.add_paragraph()
            add_text(p, sg["unvan"])
            fmt(p, WD_ALIGN_PARAGRAPH.CENTER, after=0, line=1.15)

    # Ekler
    ekler = data.get("ekler") or []
    if ekler:
        sp = doc.add_paragraph()
        fmt(sp, after=12, line=1.0, keep_next=True)
        p = doc.add_paragraph()
        r = p.add_run("Ek:")
        set_font(r, 12, bold=True)
        r.underline = True
        fmt(p, WD_ALIGN_PARAGRAPH.LEFT, after=2, keep_next=True, line=1.15)
        for i, e in enumerate(ekler):
            p = doc.add_paragraph()
            add_text(p, f"- {e}")
            fmt(p, WD_ALIGN_PARAGRAPH.LEFT, after=0, hanging=0.5,
                keep_next=(i < len(ekler) - 1), line=1.15)

    # Boş ayırıcı paragrafları kısalt
    for par in doc.paragraphs[1:]:
        if not par.text.strip() and not par.runs:
            par.paragraph_format.line_spacing = Pt(7)
            par.paragraph_format.space_after = Pt(0)

    doc.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description="SEDDK yakın izleme resmi yazısı üretir (v3)")
    ap.add_argument("girdi", help="Girdi JSON dosyası")
    ap.add_argument("cikti", help="Çıktı .docx yolu")
    default_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                               "assets", "sirket_bilgileri.json")
    ap.add_argument("--sirket", default=default_cfg, help="Şirket bilgileri JSON")
    ap.add_argument("--validate-only", action="store_true",
                    help="Sadece doğrula, dosya üretme")
    a = ap.parse_args()

    if not os.path.isfile(a.girdi):
        sys.exit(f"Girdi dosyası bulunamadı: {a.girdi}")
    if not os.path.isfile(a.sirket):
        sys.exit(f"Şirket dosyası bulunamadı: {a.sirket}")

    with open(a.girdi, encoding="utf-8") as f:
        data = json.load(f)
    with open(a.sirket, encoding="utf-8") as f:
        sirket = json.load(f)

    errors, warnings = validate_input(data, sirket)
    if warnings:
        for w in warnings:
            print(f"Uyarı: {w}", file=sys.stderr)
    if errors:
        for e in errors:
            print(f"Hata: {e}", file=sys.stderr)
        sys.exit(1)

    if a.validate_only:
        print("Doğrulama başarılı.")
        left = collect_placeholders(data, sirket)
        if left:
            print("Yer tutucular:")
            for x in left:
                print("  ", x)
        return

    out_dir = os.path.dirname(os.path.abspath(a.cikti))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    build(data, sirket, a.cikti)
    left = collect_placeholders(data, sirket)
    print(f"Yazı üretildi: {a.cikti}")
    if left:
        print("Doldurulması gereken yer tutucular:")
        for x in left:
            print("  ", x)
    else:
        print("Yer tutucu yok — yazı gönderime hazır görünüyor.")
    print("Sonraki adım: python scripts/dogrula.py " + a.cikti)


if __name__ == "__main__":
    main()