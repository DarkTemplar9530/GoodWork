import streamlit as st
from docx import Document
from docx.shared import Mm, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import re

# --- КОНСТАНТЫ ОФОРМЛЕНИЯ УрГУПС ---
FONT_NAME = 'Times New Roman'
FONT_SIZE_MAIN = Pt(14)
FONT_SIZE_HEADER = Pt(16)  # Основной + 2 пт
FONT_SIZE_TABLE = Pt(12)   # Допускается 8-14, берем 12 для компактности
INDENT_MAIN = Cm(1.25)
INDENT_NONE = Cm(0)

# Интервалы (в пунктах, приблизительно)
# 1 строка 14pt * 1.5 ≈ 21pt.
SPACE_BEFORE_SECTION = Pt(42) # 3 интервала
SPACE_AFTER_SECTION = Pt(30)  # ~2-3 интервала
SPACE_SUBSECTION = Pt(28)     # 2 интервала

def set_page_settings(doc):
    """1. Общие требования: Поля"""
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(10)
        # Отключаем связь с предыдущим для корректной нумерации
        section.footer.is_linked_to_previous = False

def clear_paragraph_format(paragraph):
    """Полная очистка форматирования параграфа"""
    p_fmt = paragraph.paragraph_format
    p_fmt.left_indent = 0
    p_fmt.right_indent = 0
    p_fmt.first_line_indent = 0
    p_fmt.space_before = 0
    p_fmt.space_after = 0
    p_fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

def apply_text_style(paragraph, size=FONT_SIZE_MAIN, bold=False, caps=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=INDENT_MAIN):
    """Применение стиля к параграфу и всем его run-ам"""
    paragraph.alignment = align
    paragraph.paragraph_format.first_line_indent = indent
    
    # Если CAPS, меняем текст
    if caps:
        text = paragraph.text.upper()
        # Аккуратно заменяем текст, стараясь сохранить структуру, если возможно, 
        # но для надежности часто проще пересоздать run-ы для заголовков
        paragraph.clear()
        paragraph.add_run(text)

    for run in paragraph.runs:
        run.font.name = FONT_NAME
        run.font.size = size
        run.font.bold = bold
        run.font.italic = False
        run.font.color.rgb = RGBColor(0, 0, 0) # Черный цвет

def add_page_number(doc):
    """4. Нумерация страниц: внизу по центру"""
    for section in doc.sections:
        footer = section.footer
        # Очищаем существующие параграфы футера
        for p in footer.paragraphs:
            p.clear()
        
        if not footer.paragraphs:
            footer.add_paragraph()
            
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = 0
        
        run = p.add_run()
        # XML для вставки поля PAGE
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        
        run.font.name = FONT_NAME
        run.font.size = FONT_SIZE_MAIN

def process_document(uploaded_file):
    doc = Document(uploaded_file)
    set_page_settings(doc)

    # Ключевые слова для структурных частей (Level 0)
    STRUCTURAL_HEADERS = [
        "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", 
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ",
        "ПРИЛОЖЕНИЕ", "ПРИЛОЖЕНИЯ", "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ"
    ]

    # Регулярные выражения
    # Разделы (1. Тема): Цифра, точка, пробел, буквы. 
    regex_level1 = re.compile(r'^\d+\.?\s+[А-Яа-яA-Za-z]') 
    # Подразделы (1.1. Тема или 1.1 Тема): Цифра, точка, цифра...
    regex_level2 = re.compile(r'^\d+(\.\d+)+\.?\s+')
    
    regex_fig = re.compile(r'^Рисунок\s+\d+', re.IGNORECASE)
    regex_tab = re.compile(r'^Таблица\s+\d+', re.IGNORECASE)

    prev_type = "text" # text, header_L0, header_L1, header_L2

    for i, para in enumerate(doc.paragraphs):
        text_raw = para.text.strip()
        if not text_raw:
            continue

        clear_paragraph_format(para)
        
        # --- 1. СТРУКТУРНЫЕ ЧАСТИ (СОДЕРЖАНИЕ, ВВЕДЕНИЕ...) ---
        # Правило: Прописные, Полужирный (+2 кегля = 16пт), По центру, Новая страница
        is_struct = False
        for key in STRUCTURAL_HEADERS:
            # Сравниваем начало строки или точное совпадение
            if text_raw.upper().startswith(key) and len(text_raw) < 100:
                is_struct = True
                break
        
        if is_struct:
            # Исключение: ПРИЛОЖЕНИЕ может быть просто словом в тексте, проверяем длину
            apply_text_style(para, size=FONT_SIZE_HEADER, bold=True, caps=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=INDENT_NONE)
            para.paragraph_format.page_break_before = True
            prev_type = "header_L0"
            continue

        # --- 2. ЗАГОЛОВКИ РАЗДЕЛОВ (1. Название) ---
        # Правило: Полужирный (+2 кегля = 16пт), С АБЗАЦНОГО ОТСТУПА (слева), Новая страница
        # Текст: С прописной буквы (не обязательно CAPS), без точки в конце
        if regex_level1.match(text_raw) and not regex_level2.match(text_raw):
            # Удаляем точку в конце, если есть
            if text_raw.endswith('.'):
                text_raw = text_raw[:-1]
                para.text = text_raw

            apply_text_style(para, size=FONT_SIZE_HEADER, bold=True, caps=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=INDENT_MAIN)
            para.paragraph_format.page_break_before = True
            
            # Отступы: сверху 3 интервала, снизу 2-3
            para.paragraph_format.space_before = SPACE_BEFORE_SECTION
            para.paragraph_format.space_after = SPACE_AFTER_SECTION
            
            prev_type = "header_L1"
            continue

        # --- 3. ПОДРАЗДЕЛЫ / ПУНКТЫ (1.1. Название) ---
        # Правило: ОБЫЧНЫЙ шрифт (не жирный, 14пт), С абзацного отступа.
        if regex_level2.match(text_raw):
             if text_raw.endswith('.'):
                text_raw = text_raw[:-1]
                para.text = text_raw
            
             apply_text_style(para, size=FONT_SIZE_MAIN, bold=False, caps=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=INDENT_MAIN)
             
             # Отступы: 2 интервала сверху и снизу
             para.paragraph_format.space_before = SPACE_SUBSECTION
             para.paragraph_format.space_after = SPACE_SUBSECTION
             
             prev_type = "header_L2"
             continue

        # --- 4. ПОДПИСИ К РИСУНКАМ ---
        # Правило: Внизу, По центру, "Рисунок Х – Название"
        if regex_fig.match(text_raw):
            if text_raw.endswith('.'):
                para.text = text_raw[:-1]
            
            apply_text_style(para, size=FONT_SIZE_MAIN, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, indent=INDENT_NONE)
            para.paragraph_format.line_spacing = 1.0 # Одинарный для подписей
            para.paragraph_format.space_before = Pt(14)
            para.paragraph_format.space_after = Pt(14)
            prev_type = "caption"
            continue

        # --- 5. ЗАГОЛОВКИ ТАБЛИЦ ---
        # Правило: Вверху, Слева (с абзацного отступа или без - в стандарте "над таблицей слева", часто трактуется как без отступа)
        if regex_tab.match(text_raw):
            if text_raw.endswith('.'):
                para.text = text_raw[:-1]
            
            # Делаем слева, без отступа
            apply_text_style(para, size=FONT_SIZE_MAIN, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, indent=INDENT_NONE)
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_before = Pt(14)
            para.paragraph_format.space_after = Pt(6)
            prev_type = "caption"
            continue

        # --- 6. ОБЫЧНЫЙ ТЕКСТ ---
        # Правило: 14пт, 1.5 интервал, 1.25 отступ, по ширине
        apply_text_style(para, size=FONT_SIZE_MAIN, bold=False, caps=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=INDENT_MAIN)
        prev_type = "text"

    # --- ОБРАБОТКА ТАБЛИЦ ---
    for table in doc.tables:
        table.autofit = False
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    # В таблицах шрифт меньше (12пт), одинарный интервал
                    p.paragraph_format.first_line_indent = INDENT_NONE
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT # Базовое, заголовки можно по центру вручную
                    
                    for run in p.runs:
                        run.font.name = FONT_NAME
                        run.font.size = FONT_SIZE_TABLE
                        # Если текст был жирным (шапка), оставляем жирным
                        if run.font.bold:
                            run.font.bold = True

    add_page_number(doc)
    return doc

# --- ИНТЕРФЕЙС ---
st.set_page_config(page_title="Нормоконтроль ВКР УрГУПС", layout="centered")

st.title("🎓 Авто-оформление ВКР (Бакалавриат УрГУПС)")
st.markdown("""
**Сервис форматирует документ согласно СТО УрГУПС 2.3.5-2022:**

1.  **Шрифты:** Times New Roman. Основной текст — 14 пт, Разделы — 16 пт (полужирный).
2.  **Заголовки:**
    *   *ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ:* По центру, прописными, 16 пт, жирный.
    *   *Разделы (1. ...):* **С абзацного отступа**, 16 пт, жирный.
    *   *Подразделы (1.1. ...):* **Обычный шрифт** (не жирный), 14 пт.
3.  **Поля:** 30 / 10 / 20 / 20 мм.
4.  **Отступы:** Абзац 1.25 см. Интервалов между абзацами нет (0 пт).
5.  **Нумерация:** Сквозная, внизу по центру.
""")

uploaded_file = st.file_uploader("Загрузите файл .docx", type="docx")

if uploaded_file is not None:
    if st.button("🛠 Применить стандарты УрГУПС"):
        with st.spinner("Обработка документа..."):
            try:
                processed_doc = process_document(uploaded_file)
                
                bio = io.BytesIO()
                processed_doc.save(bio)
                bio.seek(0)
                
                st.success("Готово! Скачайте файл ниже.")
                st.download_button(
                    label="📥 Скачать оформленную ВКР",
                    data=bio,
                    file_name=f"UrGUPS_Fixed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.info("Примечание: Титульный лист рекомендуется проверять вручную, так как он имеет сложную структуру таблиц/подписей.")
            except Exception as e:
                st.error(f"Ошибка: {e}")
