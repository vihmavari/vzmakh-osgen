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
st.set_page_config(page_title="Генератор отчетов", page_icon="🖨")
st.title("Генератор отчетов по успеваемости")

# --- БОКОВАЯ ПАНЕЛЬ НАСТРОЕК ---
st.sidebar.header("⚙️ Настройки отчета")

# Настройка фильтра по дате
st.sidebar.subheader("Фильтр данных")
start_date = st.sidebar.date_input(
    "Показывать оценки начиная с:",
    value=datetime.date(2026, 1, 1),
    help="Оценки за даты ранее выбранной не попадут в отчет"
)

# Настройка внешнего вида
st.sidebar.subheader("Внешний вид")
show_dates = st.sidebar.toggle("Включить строку дат в таблице", value=True)

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
PAGE_WIDTH_CM = 19.0  # Рабочая ширина листа A4 (21см - поля)


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

# --- 3. ЗАГРУЗКА ФАЙЛА ---
uploaded_file = st.file_uploader("Выберите Excel файл (журнал класса)", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Создать ОС"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        output_doc = BytesIO()
        results = []

        try:
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
                status_text.text(f"🔍 Считывание: {sheet} ({i + 1}/{total_sheets})")

                if sheet not in sheet_names:
                    continue

                df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

                try:
                    topics = df.iloc[0, 3:].tolist()
                    dates_raw = df.iloc[1, 3:].tolist()
                    students = df.iloc[5:, :]
                except:
                    continue

                for _, row in students.iterrows():
                    student = row[1]
                    if not isinstance(student, str) or not student.strip():
                        continue

                    grades = row[3:].tolist()
                    for topic, date_val, GRADE_VAL in zip(topics, dates_raw, grades):
                        try:
                            current_date = pd.to_datetime(date_val).date()
                            if current_date < start_date:
                                continue
                            date_fmt = current_date.strftime("%d.%m")
                        except:
                            continue

                        formatted_g = format_grade(GRADE_VAL)
                        if formatted_g == "" or formatted_g.lower() == "nan":
                            continue

                        results.append({
                            "ФИО": student.strip(),
                            "Предмет": str(sheet).strip(),
                            "Тема": str(topic).strip(),
                            "Дата": date_fmt,
                            "Оценка": formatted_g,
                        })

            # ------------------ ЭТАП 2: ГЕНЕРАЦИЯ WORD ------------------
            if not results:
                st.error(f"❌ За период с {start_date.strftime('%d.%m.%Y')} оценок не найдено.")
            else:
                with st.spinner("✍️ Оформляем Word-документ..."):
                    full = pd.DataFrame(results).sort_values(["ФИО", "Предмет", "Дата"])
                    doc = Document()
                    for sec in doc.sections:
                        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(0.9)

                    student_groups = list(full.groupby("ФИО"))
                    total_students = len(student_groups)

                    for idx, (student, df_student) in enumerate(student_groups):
                        status_text.text(f"📄 Оформление: {student} ({idx + 1}/{total_students})")
                        doc.add_heading(student, level=2)

                        # Подготовка данных предметов
                        subjects_data = []
                        for subj_name, df_subj in df_student.groupby("Предмет"):
                            t_list = df_subj["Тема"].tolist()
                            if not t_list: continue
                            subjects_data.append({
                                "name": subj_name,
                                "topics": t_list,
                                "dates": df_subj["Дата"].tolist(),
                                "grades": df_subj["Оценка"].tolist(),
                                "width_cm": len(t_list) * MAX_WIDTH_CM
                            })

                        # Упаковка в ряды
                        rows = []
                        current_row = []
                        current_row_w = 0
                        for s in subjects_data:
                            if current_row_w + s["width_cm"] > PAGE_WIDTH_CM and current_row:
                                rows.append(current_row)
                                current_row = []
                                current_row_w = 0
                            current_row.append(s)
                            current_row_w += s["width_cm"] + 1
                        if current_row: rows.append(current_row)

                        # Отрисовка
                        for row_subjects in rows:
                            container = doc.add_table(rows=1, cols=len(row_subjects))
                            container.autofit = False

                            for col_idx, s in enumerate(row_subjects):
                                parent_cell = container.rows[0].cells[col_idx]
                                p = parent_cell.paragraphs[0]
                                p.add_run(s["name"]).bold = True

                                num_rows = 3 if show_dates else 2
                                inner_table = parent_cell.add_table(rows=num_rows, cols=len(s["topics"]))
                                inner_table.style = 'Table Grid'

                                # Фиксация вложенной таблицы
                                itblPr = inner_table._tbl.tblPr
                                itblLayout = OxmlElement('w:tblLayout')
                                itblLayout.set(qn('w:type'), 'fixed')
                                itblPr.append(itblLayout)

                                for c_idx, (t, d, g) in enumerate(zip(s["topics"], s["dates"], s["grades"])):
                                    cell_t = inner_table.rows[0].cells[c_idx]
                                    cell_t.text = str(t) if str(t) != 'nan' else ''
                                    tcPr = cell_t._tc.get_or_add_tcPr()
                                    rotation = parse_xml(r'<w:textDirection {} w:val="btLr"/>'.format(nsdecls('w')))
                                    tcPr.append(rotation)

                                    if show_dates:
                                        inner_table.rows[1].cells[c_idx].text = str(d)
                                    inner_table.rows[-1].cells[c_idx].text = str(g)

                                    for r_idx in range(num_rows):
                                        tcW = inner_table.rows[r_idx].cells[c_idx]._tc.get_or_add_tcPr().get_or_add_tcW()
                                        tcW.set(qn('w:w'), str(max_col_width_dxa))
                                        tcW.set(qn('w:type'), 'dxa')

                                max_len = max(len(str(t)) for t in s["topics"]) if s["topics"] else 1
                                inner_table.rows[0].height = Cm(max(max_len * HEIGHT_COEFF, BASE_HEIGHT_CM))
                                inner_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                            doc.add_paragraph()  # Отступ между рядами предметов

                        if idx < total_students - 1:
                            if separator_type == "Разрыв страницы":
                                doc.add_page_break()
                            else:
                                for _ in range(num_paragraphs):
                                    doc.add_paragraph()

                    doc.save(output_doc)

                progress_bar.progress(1.0)
                status_text.empty()
                st.success(f"✅ Отчет успешно сформирован!")
                filename = (f"Обратная_Связь_{datetime.datetime.now().strftime('%d.%m')} "
                            f"({uploaded_file.name.split('.')[0]}).docx")
                st.download_button(
                    label="📥 Скачать обратную связь",
                    data=output_doc.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Произошла критическая ошибка: {e}")
