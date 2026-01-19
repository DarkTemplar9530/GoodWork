import streamlit as st
from docx import Document
from docx.shared import Mm, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

def set_margins(doc):
    """1. Установка полей"""
    sections = doc.sections
    for section in sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(10)

def configure_styles(doc):
    """Настройка базовых стилей (Normal)"""
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    
    # Настройка абзаца для Normal
    p_format = style.paragraph_format
    p_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p_format.first_line_indent = Cm(1.25)
    p_format.space_after = Pt(0)
    p_format.space_before = Pt(0)

def is_structural_header(text):
    """Проверка, является ли текст заголовком структурной части"""
    keywords = [
        "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", 
        "СПИСОК ИСПОЛЬЗУЕМЫХ ИСТОЧНИКОВ", "ПРАКТИЧЕСКАЯ РАБОТА"
    ]
    cleaned = text.strip().upper()
    for key in keywords:
        if key in cleaned:
            return True
    return False

def format_paragraph_text(paragraph):
    """Принудительное форматирование текста внутри абзаца"""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    paragraph.paragraph_format.first_line_indent = Cm(1.25)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    
    for run in paragraph.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)

def process_headers(paragraph):
    """Обработка заголовков (Жирный, По центру, Без отступа, Caps)"""
    text = paragraph.text.strip()
    
    if is_structural_header(text):
        # 3, 4, 6. Заголовки разделов: Прописные, Полужирный, По центру (обычно)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        
        # Очищаем старые раны и создаем новый с правильным стилем
        paragraph.clear()
        run = paragraph.add_run(text.upper())
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        return True
        
    return False

def set_page_numbering(doc):
    """
    11, 9. Нумерация страниц.
    В python-docx нет простого способа добавить автоматическую нумерацию Word (Field),
    но мы можем настроить футер, чтобы пользователь мог обновить поле.
    Ниже приводится код добавления поля {PAGE} в футер.
    """
    for section in doc.sections:
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Очистить футер
        paragraph.clear()
        
        # Добавить поле страницы (XML хак)
        run = paragraph.add_run()
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
        
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)

def process_tables(doc):
    """7, 14. Обработка таблиц"""
    for table in doc.tables:
        # Отступ сверху и снизу (вставляем пустые параграфы до и после, если их нет)
        # Технически сложно вставить параграф ПЕРЕД таблицей через python-docx без move_table
        # Поэтому мы настроим стиль текста внутри
        
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        # Допускается 12-14, ставим 12 для компактности или 14 по стандарту
                        if run.font.size is None or run.font.size > Pt(14) or run.font.size < Pt(10):
                             run.font.size = Pt(14)
                    
                    # В таблицах обычно нет красной строки и интервал 1
                    paragraph.paragraph_format.first_line_indent = Cm(0)
                    paragraph.paragraph_format.line_spacing = 1

def process_document(uploaded_file):
    doc = Document(uploaded_file)
    
    # 1. Установка полей
    set_margins(doc)
    
    # 2. Настройка базового стиля
    configure_styles(doc)
    
    # Проход по всем параграфам
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
            
        # Попытка определить заголовок
        is_header = process_headers(para)
        
        if not is_header:
            # Если это не заголовок, применяем стандартное оформление (Пункт 2)
            # Проверяем, не является ли это подписью к рисунку или таблице (по ключевым словам)
            text_lower = para.text.lower().strip()
            if text_lower.startswith("рисунок") or text_lower.startswith("таблица"):
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.first_line_indent = Cm(0)
                para.paragraph_format.line_spacing = 1 # Одинарный для подписей
                for run in para.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
                    run.font.bold = False # Обычная жирность (п. 7)
            else:
                # Обычный текст
                format_paragraph_text(para)

    # Обработка таблиц
    process_tables(doc)
    
    # Нумерация страниц
    set_page_numbering(doc)
    
    return doc

# --- Интерфейс приложения (Streamlit) ---
st.title("Автоматическое оформление по ГОСТ/Стандартам")
st.markdown("""
Этот сервис приводит оформление `.docx` файлов к следующим стандартам:
- **Поля:** 30-10-20-20 мм
- **Шрифт:** Times New Roman, 14 пт
- **Интервал:** 1.5 строки
- **Отступ:** 1.25 см
- **Заголовки:** ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и др. по центру и жирным.
""")

uploaded_file = st.file_uploader("Загрузите ваш файл .docx", type="docx")

if uploaded_file is not None:
    st.info("Файл обрабатывается...")
    
    try:
        # Обработка
        processed_doc = process_document(uploaded_file)
        
        # Сохранение в буфер памяти
        bio = io.BytesIO()
        processed_doc.save(bio)
        bio.seek(0)
        
        st.success("Готово! Скачайте отформатированный файл.")
        
        st.download_button(
            label="Скачать оформленный документ",
            data=bio,
            file_name=f"formatted_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        st.warning("""
        **Обратите внимание:** 
        Программа пытается автоматически распознать структуру, но сложные элементы (формулы, разрывы разделов, специфичные таблицы) 
        рекомендуется проверить вручную после скачивания.
        """)
        
    except Exception as e:
        st.error(f"Произошла ошибка при обработке файла: {e}")