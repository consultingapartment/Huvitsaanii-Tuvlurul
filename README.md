# ҮЦТХТ төвлөрлийн судалгаа: автомат шинэчлэл

Өдөр бүр mcsd.mn-ээс шинэ "Үнэт цаасны төвлөрлийн судалгаа" PDF-ийг шалгаж, байвал:
1. `pdfs/YYYY-MM-DD.pdf` болгон татна (`uctht_fetch.py`)
2. PDF-ийг задалж (`pdf_parser.py`, нийт дүнгээр баталгаажуулна) `data/uctht_concentration.xlsx`-д
   шинэ `Тайлан YY-MM-DD` хуудас болон 3 нэгтгэл хуудсанд багана нэмнэ (`update_workbook.py`)
3. `output/infographic_YYYY-MM-DD.png` (3240px өргөн) үүсгэнэ (`uctht_infographic.py`)
4. Үр дүнг repo-д commit хийнэ (`.github/workflows/update.yml`)

Улирлын сүүл (3, 6, 9, 12) тайлан өмнөх улирлын сүүлтэй, бусад нь өмнөх сартай харьцуулагдана.

## Тохируулах
1. Энэ хавтсыг GitHub repo болгон push хийнэ. Settings → Actions → General → Workflow permissions:
   **Read and write**.
2. Actions таб → "ҮЦТХТ төвлөрлийн тайлан шинэчлэх" → **Run workflow** гэж туршина.

## Гараар
```
pip install -r requirements.txt && python -m playwright install chromium   # + sudo apt install poppler-utils
python uctht_fetch.py --list          # линкүүдийг шалгах
python run_pipeline.py                # бүтэн урсгал
python uctht_infographic.py data/uctht_concentration.xlsx --scale 4   # зөвхөн инфографик
```
Excel нээхэд нэгтгэлийн "Өөрчлөлт %" томьёо автоматаар тооцогдоно.
Шинэ ҮЦ код гарвал нэрийг PDF-ээс авна (урт нэр таслагдах магадлалтай), лог дээр анхааруулна.
