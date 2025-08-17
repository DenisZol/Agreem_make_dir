# -*- coding: utf-8 -*-
"""
Скрипт: make_bank_letter_from_agreement_v6.py

Что нового по сравнению с v5:
- Возвращена логика работы с файлами в одной папке (без папки INPUT).
- Реализовано предложение пользователя: CASE_NUM ищется не по всей странице,
  а в определенной области вверху документа для большей надежности.
- Убраны input() для паузы в конце работы для упрощения.
- Возвращено перемещение файлов (move) вместо копирования согласно ТЗ.
"""
import os, re, shutil, datetime
from decimal import Decimal, InvalidOperation

try:
    import pdfplumber
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.document import Document as DocumentObject
    from docx.table import _Cell
except Exception as e:
    raise SystemExit(f"Не установлены зависимости: {e}")

# Предварительно скомпилированные регулярные выражения
CASE_NUM_PATTERN = re.compile(r"\b0+(\d{5,})\b")
CASE_NUM_FALLBACK_PATTERN = re.compile(r"\b\d{6,9}\b")
AMOUNT_PATTERNS = [
    re.compile(r"(?:amount of|USD|\$)\s*([0-9][0-9 ,.]*)", flags=re.IGNORECASE),
    re.compile(r"USD\s*\$?\s*([0-9][0-9 ,.]*)", flags=re.IGNORECASE),
]
NON_DIGITS_PATTERN = re.compile(r"[^\d\.]")
UA_PURPOSE_PATTERN = re.compile(r"у вигляді\s+([^.]+)\.", flags=re.IGNORECASE)
DATE_US_PATTERN = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
GRANT_AGREEMENT_PATTERN = re.compile(r"Grant Agreement.*\.pdf", flags=re.IGNORECASE)

# ... (функции ua_date, with_thin_space_groups, find_amount, find_ua_purpose, max_date_us остаются без изменений) ...

UA_MONTHS_GEN = {
    1: "січня", 2: "лютого", 3: "березня", 4: "квітня",
    5: "травня", 6: "червня", 7: "липня", 8: "серпня",
    9: "вересня", 10: "жовтня", 11: "листопада", 12: "грудня",
}

def ua_date(dt: datetime.date) -> str:
    dd = f"{dt.day:02d}"
    return f"«{dd}» {UA_MONTHS_GEN[dt.month]} {dt.year} року"

def with_thin_space_groups(n: Decimal) -> str:
    return f"{n:,.2f}".replace(",", " ").replace(".", ".")

def find_case_num_in_crop(page) -> str | None:
    """Ищет номер дела в верхней части страницы для надежности."""
    # Обрезаем верхнюю часть страницы (первые 100 пикселей по высоте)
    top_of_page = page.crop((0, 0, page.width, 100))
    text = top_of_page.extract_text(x_tolerance=1) or ""
    
    # Ищем номер, который начинается с нулей
    m = CASE_NUM_PATTERN.search(text)
    if m:
        return m.group(1) # Возвращаем число без ведущих нулей
    # Резервный поиск, если не нашли по основному паттерну
    m2 = CASE_NUM_FALLBACK_PATTERN.search(text)
    return str(int(m2.group(0))) if m2 else None

def find_amount(text_first_page: str) -> Decimal | None:
    for pat in AMOUNT_PATTERNS:
        for m in pat.finditer(text_first_page):
            digits = NON_DIGITS_PATTERN.sub("", m.group(1))
            try:
                return Decimal(digits)
            except InvalidOperation:
                continue
    return None

def find_ua_purpose(full_text: str) -> str | None:
    m = UA_PURPOSE_PATTERN.search(full_text)
    return ("у вигляді " + m.group(1).strip()) if m else None

def max_date_us(text_last_page: str) -> datetime.date | None:
    dates = []
    for m in DATE_US_PATTERN.finditer(text_last_page):
        mm, dd, yyyy = map(int, m.groups())
        try:
            dates.append(datetime.date(yyyy, mm, dd))
        except ValueError:
            pass
    return max(dates) if dates else None

# ... (функции _replace_in_block, replace_placeholders остаются без изменений) ...

def _replace_in_paragraph(paragraph: Paragraph, placeholders: dict):
    """Replace placeholders in a paragraph without recursion.

    All placeholders are processed in the order provided by ``placeholders``.
    The function correctly handles cases where a placeholder spans across
    several ``Run`` objects within the paragraph."""

    while True:
        full_text = "".join(run.text for run in paragraph.runs)
        replaced = False

        for key, value in placeholders.items():
            start_index = full_text.find(key)
            if start_index == -1:
                continue

            end_index = start_index + len(key)
            runs_to_modify = []
            current_pos = 0
            for run in paragraph.runs:
                run_len = len(run.text)
                if start_index < current_pos + run_len and current_pos < end_index:
                    runs_to_modify.append(run)
                current_pos += run_len

            if not runs_to_modify:
                continue

            first_run = runs_to_modify[0]
            last_run = runs_to_modify[-1]

            prefix = suffix = ""
            current_pos = 0
            for run in paragraph.runs:
                run_len = len(run.text)
                if run is first_run:
                    prefix = run.text[:start_index - current_pos]
                if run is last_run:
                    suffix = run.text[end_index - current_pos:]
                    break
                current_pos += run_len

            first_run.text = prefix + str(value)
            for r in runs_to_modify[1:]:
                r.text = ""

            if first_run is last_run:
                first_run.text += suffix
            else:
                last_run.text = suffix

            replaced = True
            break

        if not replaced:
            break

def _replace_in_block(block, placeholders: dict):
    for p in getattr(block, "paragraphs", []):
        _replace_in_paragraph(p, placeholders)
    for table in getattr(block, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                _replace_in_block(cell, placeholders)

def replace_placeholders(doc: DocumentObject, placeholders: dict):
    _replace_in_block(doc, placeholders)
    for section in doc.sections:
        _replace_in_block(section.header, placeholders)
        _replace_in_block(section.footer, placeholders)

def process_file(pdf_path: str, template_path: str):
    base_dir = os.path.dirname(pdf_path)
    print(f"🔎 Обрабатываю: {os.path.basename(pdf_path)}")

    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        # Извлекаем CASE_NUM из верхней части страницы, как вы и предложили
        case_num = find_case_num_in_crop(first_page)
        
        # Остальной текст для других данных
        first_page_text = first_page.extract_text(x_tolerance=1) or ""
        last_page_text = pdf.pages[-1].extract_text(x_tolerance=1) or ""
        full_text = "\n".join((p.extract_text(x_tolerance=1) or "") for p in pdf.pages)

    # Ищем остальные данные в тексте
    amount = find_amount(first_page_text)
    doc_date = max_date_us(last_page_text)
    case_descr = find_ua_purpose(full_text)

    errors = []
    if not case_num:
        errors.append("не найден CASE_NUM")
    if amount is None:
        errors.append("не найдена сумма (FULL_AMOUNT)")
    if not doc_date:
        errors.append("не найдена дата (DATE)")
    if not case_descr:
        print("⚠️  CASE_DESCR не найден — поле будет пустым (в текущем шаблоне не используется).")

    if errors:
        print("⛔ " + "; ".join(errors) + f" — пропускаю файл {os.path.basename(pdf_path)}")
        print()
        return

    YY_MM = f"{doc_date.year % 100:02d}-{doc_date.month:02d}"
    FULL_AMOUNT_DEC = f"{int(amount):d}"
    FULL_AMOUNT_fmt = with_thin_space_groups(Decimal(f"{amount:.2f}"))

    DATE_uA = ua_date(doc_date)
    DATE_plus_2 = ua_date(doc_date + datetime.timedelta(days=2))
    DATE_plus_3 = ua_date(doc_date + datetime.timedelta(days=3))
    DATE_MM_ONLY = f"{doc_date.month:02d}"

    folder_name = f"{YY_MM} Нова ХХХ {FULL_AMOUNT_DEC} №{case_num} Хелп"
    out_dir = os.path.join(base_dir, folder_name)

    if os.path.exists(out_dir):
        print(f"↪️ Папка уже существует: {folder_name} — пропускаю.")
        print()
        return

    os.makedirs(out_dir, exist_ok=True)

    doc = Document(template_path)
    placeholders = {
        "{{CASE_NUM}}": case_num,
        "{{FULL_AMOUNT_DEC}}": FULL_AMOUNT_DEC,
        "{{FULL_AMOUNT}}": str(FULL_AMOUNT_fmt),
        "{{DATE}}": DATE_uA,
        "{{DATE+2}}": DATE_plus_2,
        "{{DATE + 2}}": DATE_plus_2,
        "{{DATE+3}}": DATE_plus_3,
        "{{DATE + 3}}": DATE_plus_3,
        "{{DATE_MM_ONLY}}": DATE_MM_ONLY,
        "{{CASE_DESCR}}": case_descr or "",
    }
    replace_placeholders(doc, placeholders)

    out_docx = os.path.join(out_dir, f"Письмо_в_банк_№{case_num}.docx")
    if os.path.exists(out_docx):
        print(f"↪️ Word уже существует: {os.path.basename(out_docx)} — пропускаю.")
    else:
        doc.save(out_docx)
        print(f"✅ Сгенерирован документ: {os.path.basename(out_docx)}")

    dst_pdf = os.path.join(out_dir, os.path.basename(pdf_path))
    if os.path.exists(dst_pdf):
        print(f"↪️ PDF уже на месте: {os.path.basename(dst_pdf)} — пропускаю перенос.")
    else:
        shutil.move(pdf_path, dst_pdf) # Возвращено перемещение файла
        print(f"📦 Перемещён PDF в: {folder_name}")

    print()

def main():
    cwd = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(cwd, "Письмо на Банк шаблон.docx")

    if not os.path.exists(template_path):
        print("⛔ Не найден шаблон 'Письмо на Банк шаблон.docx' рядом со скриптом.")
        return

    files = [f for f in os.listdir(cwd) if GRANT_AGREEMENT_PATTERN.fullmatch(f)]
    if not files:
        print("ℹ️ Нет файлов по маске 'Grant Agreement*.pdf' в папке со скриптом.")
        return

    print(f"Найдено PDF: {len(files)}")
    for name in files:
        try:
            process_file(os.path.join(cwd, name), template_path)
        except Exception as e:
            print(f"💥 Ошибка при обработке {name}: {e}\n— Перехожу к следующему файлу.")
            print()

if __name__ == "__main__":
    main()

