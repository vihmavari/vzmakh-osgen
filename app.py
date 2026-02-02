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

# НАСТРОЙКИ ТАБЛИЦЫ
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
        # Создаем контейнеры для прогресса, чтобы они были в начале страницы
        progress_bar = st.progress(0)
        status_text = st.empty()

        output_doc = BytesIO()
        results = []

        # Читаем список листов
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        index_sheet_name = sheet_names[0]
        index_df = pd.read_excel(uploaded_file, sheet_name=index_sheet_name, header=None)
        subject_sheets = index_df.iloc[:, 0].dropna().tolist()

        total_sheets = len(subject_sheets)

        # ------------------ ЭТАП 1: СБОР ДАННЫХ ------------------
        for i, sheet in enumerate(subject_sheets):
            # Обновляем индикатор
            progress = (i) / total_sheets
            progress_bar.progress(progress)
            status_text.text(f"🔍 Считывание данных: предмет '{sheet}' ({i + 1}/{total_sheets})")

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
            except Exception as e:
                st.warning(f"Пропущен лист '{sheet}': ошибка формата")

        # ------------------ ЭТАП 2: ГЕНЕРАЦИЯ WORD ------------------
        if not results:
            st.error("Данные не найдены. Проверьте формат Excel-файла.")
        else:
            status_text.text("📝 Формирование Word-документа (это может занять время)...")
            progress_bar.progress(0.9)  # Почти готово

            with st.spinner("Рисуем таблицы и настраиваем стили..."):
                full = pd.DataFrame(results).sort_values(["ФИО", "Предмет", "Дата"])
                doc = Document()

                # Настройка полей
                section = doc.sections[0]
                section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(1)

                counter = 0
                student_groups = full.groupby("ФИО")
                total_students = len(student_groups)

                for student, df_student in student_groups:
                    counter += 1
                    # Обновляем текст статуса для каждого ученика (опционально)
                    status_text.text(f"📄 Оформление: {student} ({counter}/{total_students})")

                    doc.add_heading(student, level=2)

                    for subject, df_subject in df_student.groupby("Предмет"):
                        doc.add_heading(subject, level=3)

                        topics = df_subject["Тема"].tolist()
                        dates = df_subject["Дата"].tolist()
                        grades = df_subject["Оценка"].tolist()

                        if not topics: continue

                        table = doc.add_table(rows=3, cols=len(topics))
                        table.style = 'Table Grid'

                        # Фиксация верстки
                        tblPr = table._tbl.tblPr
                        tblLayout = OxmlElement('w:tblLayout')
                        tblLayout.set(qn('w:type'), 'fixed')
                        tblPr.append(tblLayout)

                        for i, t in enumerate(topics):
                            cell = table.rows[0].cells[i]
                            cell.text = str(t) if str(t) != 'nan' else ''
                            tcPr = cell._tc.get_or_add_tcPr()
                            rotation = parse_xml(r'<w:textDirection {} w:val="btLr"/>'.format(nsdecls('w')))
                            tcPr.append(rotation)

                        for i, d in enumerate(dates):
                            table.rows[1].cells[i].text = str(d)
                        for i, g in enumerate(grades):
                            table.rows[2].cells[i].text = str(g)

                        # Настройка размеров
                        max_len = max(len(str(t)) for t in topics) if topics else 1
                        table.rows[0].height = Cm(max(max_len * HEIGHT_COEFF, BASE_HEIGHT_CM))
                        table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                        for col_idx in range(len(topics)):
                            for r in table.rows:
                                tcW = r.cells[col_idx]._tc.get_or_add_tcPr().get_or_add_tcW()
                                tcW.set(qn('w:w'), str(max_col_width_dxa))
                                tcW.set(qn('w:type'), 'dxa')

                    doc.add_page_break()

                doc.save(output_doc)

            # Финализация интерфейса
            progress_bar.progress(1.0)
            status_text.empty()
            st.success(f"✅ Готово! Обработано учеников: {counter}")

            st.download_button(
                label="📥 Скачать готовый документ",
                data=output_doc.getvalue(),
                file_name=f"Успеваемость_{datetime.datetime.now().strftime('%d_%m_%Y')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )