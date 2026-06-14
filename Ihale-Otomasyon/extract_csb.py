"""CSB birim fiyat listesi PDF'den poz numaralarını çıkarır."""
from __future__ import annotations

import logging
import re
import sys
from pathlib import Path

from logger_config import setup_logger

logger = setup_logger("ihale_otomasyonu")

PDF_PATH = Path(__file__).parent / "pdfler" / "CSB İNŞAAT BİRİM FİYAT LİSTESİ.pdf"
OUT_XLSX = Path(__file__).parent / "csb_poz_numaralari.xlsx"

# CSB 2026 listesi: 15.100.1001 (XX.XXX.XXXX)
POZ_RE = re.compile(r"\b(\d{2}\.\d{3}\.\d{4})\b")
ROW_RE = re.compile(
    r"^(\d{2}\.\d{3}\.\d{4})\s+(.+?)\s+"
    r"((?:\d+\s+)?(?:m[²³]?|m³|Ton|Ad|kg|g|lt|1000\s+Ad|100\s+m[²³]?|m²|m³|m|t|ad)\.?)\s+"
    r"([\d.,]+)\s*$",
    re.I | re.M,
)


def extract_text_pymupdf(pdf: Path) -> str:
    try:
        import fitz
        parts: list[str] = []
        doc = fitz.open(pdf)
        for page in doc:
            parts.append(page.get_text())
        doc.close()
        logger.info(f"PDF metni çıkarıldı (PyMuPDF): {pdf.name} ({len(''.join(parts)):,} karakter)")
        return "\n".join(parts)
    except Exception as e:
        logger.error(f"PyMuPDF ile PDF okunamadı: {pdf} - {e}")
        raise


def extract_text_pypdf(pdf: Path) -> str:
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(pdf))
        text = "\n".join(p.extract_text() or "" for p in reader.pages)
        logger.info(f"PDF metni çıkarıldı (PyPDF): {pdf.name} ({len(text):,} karakter)")
        return text
    except Exception as e:
        logger.error(f"PyPDF ile PDF okunamadı: {pdf} - {e}")
        raise


def find_poz_numbers(text: str) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for m in POZ_RE.finditer(text):
        poz = m.group(1)
        if poz not in seen:
            seen.add(poz)
            ordered.append(poz)
    return ordered


def parse_rows_simple(text: str) -> list[dict[str, str]]:
    """Satır satır: poz satırından sonraki tanım/birim/fiyat bloklarını birleştirir."""
    lines = [ln.strip() for ln in text.splitlines()]
    rows: list[dict[str, str]] = []
    i = 0
    while i < len(lines):
        m = POZ_RE.match(lines[i])
        if not m:
            i += 1
            continue
        poz = m.group(1)
        rest = lines[i][m.end() :].strip()
        i += 1
        block: list[str] = [rest] if rest else []
        while i < len(lines) and not POZ_RE.match(lines[i]):
            if lines[i]:
                block.append(lines[i])
            i += 1
        joined = " ".join(block)
        unit_m = re.search(
            r"\b((?:\d+\s+)?(?:1000\s+Ad|100\s+m[²³]?|m[²³]?|m³|Ton|Ad|kg|lt|g|t|ad))\.?\s+([\d.,]+)\s*$",
            joined,
            re.I,
        )
        if unit_m:
            tanim = joined[: unit_m.start()].strip()
            birim = unit_m.group(1).strip()
            fiyat = unit_m.group(2).strip()
        else:
            tanim, birim, fiyat = joined, "", ""
        rows.append({"poz_no": poz, "tanim": tanim, "birim": birim, "fiyat": fiyat})
    return rows


def write_excel(poz_list: list[str], table: list[dict[str, str]], out: Path) -> None:
    try:
        from openpyxl import Workbook

        out.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "CSB Poz Listesi"
        ws.append(["Sıra", "Poz No", "Tanım", "Birim", "Fiyat (TL)"])
        if table:
            for i, r in enumerate(table, start=1):
                ws.append([i, r["poz_no"], r["tanim"], r["birim"], r["fiyat"]])
        else:
            for i, poz in enumerate(poz_list, start=1):
                ws.append([i, poz, "", "", ""])
        wb.save(out)
        logger.info(f"Excel dosyası kaydedildi: {out}")
    except Exception as e:
        logger.error(f"Excel dosyası kaydedilemedi: {out} - {e}")
        raise


def main() -> int:
    from logger_config import log_program_start, log_program_end, log_summary
    
    try:
        log_program_start(logger)
        
        if not PDF_PATH.is_file():
            logger.error(f"PDF bulunamadı: {PDF_PATH}")
            print(f"[HATA] PDF bulunamadı: {PDF_PATH}", file=sys.stderr)
            log_program_end(logger, 1)
            return 1

        try:
            text = extract_text_pymupdf(PDF_PATH)
        except ImportError:
            logger.warning("PyMuPDF kütüphanesi bulunamadı, PyPDF denenecek")
            try:
                text = extract_text_pypdf(PDF_PATH)
            except ImportError as e:
                logger.error("Her iki PDF kütüphanesi de bulunamadı")
                print(f"[HATA] Hata: {e}", file=sys.stderr)
                log_program_end(logger, 1)
                return 1
        except Exception as e:
            logger.error(f"PDF okunamadı: {e}")
            print(f"[HATA] Hata: {e}", file=sys.stderr)
            log_program_end(logger, 1)
            return 1

        poz_list = find_poz_numbers(text)
        table = parse_rows_simple(text)
        write_excel(poz_list, table, OUT_XLSX)

        summary = {
            "Sayfa metni": f"{len(text):,} karakter",
            "Bulunan poz sayısı": len(poz_list),
            "Excel dosyası": str(OUT_XLSX),
        }
        log_summary(logger, summary)

        print(f"[OK] Sayfa metni: {len(text):,} karakter")
        print(f"[OK] Bulunan poz sayısı: {len(poz_list)}")
        print(f"[OK] Excel: {OUT_XLSX}")
        if poz_list[:15]:
            print(f"[OK] İlk 15: {', '.join(poz_list[:15])}")
        
        log_program_end(logger, 0)
        return 0
    except Exception as e:
        logger.exception(f"Beklenmeyen hata: {e}")
        print(f"[HATA] Beklenmeyen hata: {e}", file=sys.stderr)
        log_program_end(logger, 1)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
