import datetime
import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE
from io import BytesIO

# --- ИНТЕРФЕЙС ---
st.set_page_config(page_title="Генератор отчетов", page_icon="📊")
st.title("Генератор отчетов по успеваемости")

# --- БОКОВАЯ ПАНЕЛЬ НАСТРОЕК ---
st.sidebar.header("Настройки отчета")

# 1. Настройка строки дат
show_dates = st.sidebar.toggle("Включить строку дат", value=True)

# 2. Настройка разделителя
separator_type = st.sidebar.selectbox(
    "Разделитель между учениками",
    ["Разрыв страницы", "Пустые строки (параграфы)"]
)

num_paragraphs = 1
if separator_type == "Пустые строки (параграфы)":
    num_paragraphs = st.sidebar.number_input("Количество пустых строк", min_value=1, max_value=10, value=2)

# КОНСТАНТЫ ТАБЛИЦЫ
MAX_WIDTH_CM = 2
BASE_HEIGHT_CM = 1.5
HEIGHT_COEFF = 0.07


def cm_to_dxa(cm):
    inches = cm / 2.54
    points = inches * 72
    return int(round(points * 20))


def format_grade(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%d/%m").lstrip("0").replace("/0", "/")
    return str(val).strip()


max_col_width_dxa = cm_to_dxa(MAX_WIDTH_CM)

uploaded_file = st.file_uploader("Выберите Excel файл", type=["xlsx"])

if uploaded_file:
    if st.button("Создать ОС"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        output_doc = BytesIO()
        results = []

        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        index_sheet_name = sheet_names[0]
        index_df = pd.read_excel(uploaded_file, sheet_name=index_sheet_name, header=None)
        subject_sheets = index_df.iloc[:, 0].dropna().tolist()

        total_sheets = len(subject_sheets)

        # ------------------ ЭТАП 1: СБОР ДАННЫХ ------------------
        for i, sheet in enumerate(subject_sheets):
            progress = (i) / total_sheets
            progress_bar.progress(progress)
            status_text.text(f"🔍 Считывание данных: {sheet} ({i + 1}/{total_sheets})")

            try:
                df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                topics = df.iloc[0, 3:].tolist()
                dates = df.iloc[1, 3:].tolist()
                students = df.iloc[5:, :]

                for _, row in students.iterrows():
                    student = row[1]
                    if not isinstance(student, str) or not student.strip():
                        continue
                    grades = row[3:].tolist()
                    for topic, date, GRADE_VAL in zip(topics, dates, grades):
                        formatted_g = format_grade(GRADE_VAL)
                        if formatted_g == "" or formatted_g.lower() == "nan":
                            continue

                        try:
                            date_fmt = pd.to_datetime(date).strftime("%d.%m")
                        except:
                            date_fmt = str(date) if not pd.isna(date) else ""

                        results.append({
                            "ФИО": student.strip(),
                            "Предмет": str(sheet).strip(),
                            "Тема": str(topic).strip(),
                            "Дата": date_fmt,
                            "Оценка": formatted_g,
                        })
            except:
                continue

        # ------------------ ЭТАП 2: ГЕНЕРАЦИЯ WORD ------------------
        if not results:
            st.error("Данные не найдены.")
        else:
            with st.spinner("Формирование документа..."):
                full = pd.DataFrame(results).sort_values(["ФИО", "Предмет", "Дата"])
                doc = Document()
                for sec in doc.sections:
                    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1)

                student_groups = list(full.groupby("ФИО"))
                total_students = len(student_groups)

                for idx, (student, df_student) in enumerate(student_groups):
                    status_text.text(f"📄 Оформление: {student} ({idx + 1}/{total_students})")
                    doc.add_heading(student, level=2)

                    for subject, df_subject in df_student.groupby("Предмет"):
                        doc.add_heading(subject, level=3)

                        t_list = df_subject["Тема"].tolist()
                        d_list = df_subject["Дата"].tolist()
                        g_list = df_subject["Оценка"].tolist()

                        if not t_list: continue

                        # Определяем количество строк и создаем таблицу
                        num_rows = 3 if show_dates else 2
                        ncols = len(t_list)
                        table = doc.add_table(rows=num_rows, cols=ncols)
                        table.style = 'Table Grid'

                        # --- ЖЕСТКАЯ ФИКСАЦИЯ ШИРИНЫ (tblLayout: fixed) ---
                        tblPr = table._tbl.tblPr
                        tblLayout = OxmlElement('w:tblLayout')
                        tblLayout.set(qn('w:type'), 'fixed')
                        tblPr.append(tblLayout)

                        row_topics = table.rows[0]
                        row_grades = table.rows[-1]
                        row_dates = table.rows[1] if show_dates else None

                        for i, (t, d, g) in enumerate(zip(t_list, d_list, g_list)):
                            # 1. Заполняем Темы (Вертикально)
                            cell_t = row_topics.cells[i]
                            cell_t.text = str(t) if str(t) != 'nan' else ''
                            tcPr = cell_t._tc.get_or_add_tcPr()
                            rotation = parse_xml(r'<w:textDirection {} w:val="btLr"/>'.format(nsdecls('w')))
                            tcPr.append(rotation)

                            # 2. Заполняем Даты (если нужно)
                            if show_dates:
                                row_dates.cells[i].text = str(d)

                            # 3. Заполняем Оценки
                            row_grades.cells[i].text = str(g)

                            # --- ВОЗВРАЩАЕМ ФИКСАЦИЮ ШИРИНЫ КОЛОНОК ---
                            for r_idx in range(num_rows):
                                cell = table.rows[r_idx].cells[i]
                                tcW = cell._tc.get_or_add_tcPr().get_or_add_tcW()
                                tcW.set(qn('w:w'), str(max_col_width_dxa))
                                tcW.set(qn('w:type'), 'dxa')

                        # Высота первой строки (шапки)
                        max_len = max(len(str(t)) for t in t_list) if t_list else 1
                        row_topics.height = Cm(max(max_len * HEIGHT_COEFF, BASE_HEIGHT_CM))
                        row_topics.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                    # --- РАЗДЕЛИТЕЛЬ МЕЖДУ УЧЕНИКАМИ ---
                    # Не добавляем разделитель после последнего ученика
                    if idx < total_students - 1:
                        if separator_type == "Разрыв страницы":
                            doc.add_page_break()
                        else:
                            for _ in range(num_paragraphs):
                                doc.add_paragraph()

                doc.save(output_doc)

            progress_bar.progress(1.0)
            status_text.empty()
            st.success(f"✅ Документ готов!")
            st.download_button(
                label="📥 Скачать результат",
                data=output_doc.getvalue(),
                file_name=f"ОС_{datetime.datetime.now().strftime('%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )