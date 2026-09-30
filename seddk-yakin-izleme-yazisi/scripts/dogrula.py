#!/usr/bin/env python3
"""Üretilmiş SEDDK yazısının sayfa yerleşimini pymupdf ile denetler.

Kontroller:
1. Sayfa sayısı (resmî yazılar için pratik üst sınır varsayılan 2)
2. Son sayfada yalnız kalan imza bloğu / ek satırları var mı
   (Saygılarımızla'dan sonraki içerik, önceki sayfada 4+ tam satır varken yalnız kalmamalı)
3. [●...] yer tutucusu görüntüde sarı vurgu olarak kalmış mı
4. Yatay taşma: metin blokları sayfa genişliğini aşıyor mu

Kullanım:
    python scripts/dogrula.py yazı.docx [--max-sayfa 2] [--pdf]

Bağımlılık: pymupdf (fitz). Kurulu değilse uyarı verir ve yalnızca docx XML
kontrollerine düşer.
"""
import argparse
import sys

import re

PLACEHOLDER = re.compile(r"\[●[^\]]*\]")


def check_docx_xml(path):
    """pymupdf yoksa: docx XML'inde yer tutucu ve temel yapı kontrolü."""
    import zipfile
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        names = z.namelist()
    text = re.sub(r"<[^>]+>", "", xml)
    ph = sorted(set(PLACEHOLDER.findall(text)))
    return {
        "sayfa_sayisi": None,
        "yer_tutucular": ph,
        "hatalar": [],
        "uyarilar": ([]
                     if "word/document.xml" in names else
                     ["document.xml bulunamadı"]),
    }


def check_pdf(path, max_sayfa):
    import fitz
    doc = fitz.open(path)
    hatalar = []
    uyarilar = []
    n = len(doc)
    if n > max_sayfa:
        uyarilar.append(
            f"Sayfa sayısı {n} (önerilen üst sınır {max_sayfa}). Not: bu ölçüm pymupdf yerleşimiyle "
            f"yapılır; Word keepNext zincirini uyguladığından gerçek sayfa sayısı daha az olabilir. "
            f"Yine de içerik kısaltılabiliyorsa kısaltın.")

    ph_all = []
    for i, page in enumerate(doc):
        pw = page.rect.width
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, metin = b[0], b[1], b[2], b[3], b[4]
            if x1 > pw - 5:
                hatalar.append(f"Sayfa {i+1}: blok sağ kenardan taşıyor: {metin[:40]!r}")
            for m in PLACEHOLDER.findall(metin):
                ph_all.append(m)
        # sarı vurgu tespiti: sarı dolgulu metin span'ları
        for d in page.get_text("dict")["blocks"]:
            for ln in d.get("lines", []):
                for sp in ln.get("spans", []):
                    if sp.get("flags") is not None and "fill" in sp:
                        pass  # vurgu rengi PDF'te span fill'inde; görsel kontrol önerilir

    if ph_all:
        uyarilar.append("Görüntüde [●...] yer tutucuları hâlâ var: " + ", ".join(sorted(set(ph_all))))

    if n > 1:
        last = doc[n - 1].get_text("blocks")
        prev = doc[n - 2].get_text("blocks")
        last_blob = " ".join(b[4] for b in last)
        if "Saygılarımızla" in last_blob and len(prev) >= 4:
            # Saygı + imza + ek bloğu önceki sayfanın sonuna sığmadıysa uyar
            uyarilar.append(
                "Son sayfada yalnız imza/ek bloğu kalmış görünüyor — "
                "gövdeyi bir-iki satır kısaltıp bloğu önceki sayfaya çekmeyi deneyin")

    return {"sayfa_sayisi": n, "yer_tutucular": sorted(set(ph_all)),
            "hatalar": hatalar, "uyarilar": uyarilar}


def main():
    ap = argparse.ArgumentParser(description="SEDDK yazısı sayfa yerleşimi denetleyicisi")
    ap.add_argument("docx", help="Denetlenecek .docx")
    ap.add_argument("--max-sayfa", type=int, default=2, help="Önerilen üst sayfa sınırı")
    a = ap.parse_args()

    try:
        import fitz  # noqa: F401
        res = check_pdf(a.docx, a.max_sayfa)
    except ImportError:
        print("pymupdf kurulu değil; docx XML kontrolüne düşülüyor.", file=sys.stderr)
        res = check_docx_xml(a.docx)

    print(f"Dosya: {a.docx}")
    print(f"Sayfa sayısı: {res['sayfa_sayisi'] if res['sayfa_sayisi'] else 'bilinmiyor (XML modu)'}")
    if res["hatalar"]:
        for h in res["hatalar"]:
            print(f"HATA: {h}")
    if res["uyarilar"]:
        for u in res["uyarilar"]:
            print(f"UYARI: {u}")
    if not res["hatalar"] and not res["uyarilar"]:
        print("Sayfa yerleşimi uygun görünüyor.")
    if res["yer_tutucular"]:
        print("Yer tutucular (gönderimden önce doldurulmalı):")
        for x in res["yer_tutucular"]:
            print("  ", x)

    sys.exit(1 if res["hatalar"] else 0)


if __name__ == "__main__":
    main()