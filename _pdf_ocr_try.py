"""Попытка читать PDF бюллетени через PyMuPDF + конвертация изображений."""
import sys
from pathlib import Path

BASE = Path(r"c:\Users\D.Muldabaev\Desktop\Приложения для сервиса")
BULL = BASE / "бюллетени"

import fitz  # PyMuPDF

IMG_DIR = BASE / "_pdf_images"
IMG_DIR.mkdir(exist_ok=True)

for pdf_path in sorted(BULL.glob("*.pdf")):
    print(f"\n=== {pdf_path.name} ===")
    try:
        doc = fitz.open(pdf_path)
        print(f"    Страниц: {len(doc)}")
        for i, page in enumerate(doc):
            text = page.get_text("text").strip()
            if text:
                print(f"  [стр.{i+1}] ТЕКСТ НАЙДЕН ({len(text)} симв.):")
                for line in text.split("\n")[:80]:
                    if line.strip():
                        print("   ", line)
            else:
                print(f"  [стр.{i+1}] Текста нет — рендерим изображение")
                mat = fitz.Matrix(2, 2)  # 2x zoom
                pix = page.get_pixmap(matrix=mat)
                out_img = IMG_DIR / f"{pdf_path.stem}_p{i+1}.png"
                pix.save(out_img)
                print(f"           Сохранено: {out_img.name}")
    except Exception as e:
        print(f"  ОШИБКА: {e}")

print("\nГотово.")
