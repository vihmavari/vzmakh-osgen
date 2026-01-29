import datetime

import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE
from io import BytesIO

st.title("Генератор отчетов по успеваемости")

# НАСТРОЙКИ
MAX_WIDTH_CM = 2
BASE_HEIGHT_CM = 1.5
HEIGHT_COEFF = 0.07


def cm_to_dxa(cm):
    inches = cm / 2.54
    points = inches * 72
    dxa = int(round(points * 20))
    return dxa


max_col_width_dxa = cm_to_dxa(MAX_WIDTH_CM)

uploaded_file = st.file_uploader("Выберите Excel файл", type=["xlsx"])

if uploaded_file:
    if st.button("Создать ОС"):
        # 1. Оборачиваем ваш скрипт в функцию, которая принимает uploaded_file
        # 2. Вместо сохранения в файл используем BytesIO
        output_doc = BytesIO()

        # --- Здесь вставляется ваш код обработки ---
        results = []
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # Первый лист содержит перечень листов с предметами/группами
        index_sheet_name = sheet_names[0]
        index_df = pd.read_excel(uploaded_file, sheet_name=index_sheet_name, header=None)

        # Предположим, что листы перечислены в первом столбце
        subject_sheets = index_df.iloc[:, 0].dropna().tolist()

        # ------------------ собираем оценки ------------------
        for sheet in subject_sheets:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
            except Exception as e:
                print(f"Ошибка при чтении листа {sheet}: {e}")
                continue

            # Темы в 1-й строке, даты во 2-й, ученики с 6-й (индекс 5)
            try:
                topics = df.iloc[0, 3:].tolist()
                dates = df.iloc[1, 3:].tolist()
                students = df.iloc[5:, :]
            except Exception as e:
                print(f"Неожиданный формат в листе {sheet}: {e}")
                continue

            for _, row in students.iterrows():
                student = row[1]
                if not isinstance(student, str) or not student.strip():
                    continue
                grades = row[3:].tolist()
                for topic, date, GRADE_VAL in zip(topics, dates, grades):
                    if pd.isna(GRADE_VAL):
                        continue
                    try:
                        date_fmt = pd.to_datetime(date).strftime("%d.%m")
                    except Exception:
                        date_fmt = str(date) if not pd.isna(date) else ""
                    results.append({
                        "ФИО": student.strip(),
                        "Предмет": str(sheet).strip(),
                        "Тема": str(topic).strip(),
                        "Дата": date_fmt,
                        "Оценка": str(GRADE_VAL).strip(),
                    })

        # ------------------ создаём DataFrame ------------------
        full = pd.DataFrame(results)
        if full.empty:
            print("Нет собранных оценок.")
            raise SystemExit

        full = full.sort_values(["ФИО", "Предмет", "Дата"])

        # ------------------ создаём Word ------------------
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1)
        section.right_margin = Cm(1)

        counter = 0
        for student, df_student in full.groupby("ФИО"):
            counter += 1
            doc.add_heading(student, level=2)

            for subject, df_subject in df_student.groupby("Предмет"):
                doc.add_heading(subject, level=3)

                topics_raw = df_subject["Тема"].tolist()
                dates_raw = df_subject["Дата"].tolist()
                grades_raw = df_subject["Оценка"].tolist()

                topics, dates, grades = [], [], []
                for t, d, g in zip(topics_raw, dates_raw, grades_raw):
                    val = str(g).strip().lower()
                    if val == "" or val == "nan":
                        continue
                    topics.append(t)
                    dates.append(d)
                    grades.append(g)

                if not topics:
                    continue

                ncols = max(1, len(topics))
                table = doc.add_table(rows=3, cols=ncols)
                table.style = 'Table Grid'
                tbl = table._tbl
                tblPr = tbl.tblPr
                tblLayout = OxmlElement('w:tblLayout')
                tblLayout.set(qn('w:type'), 'fixed')
                existing = tblPr.find(qn('w:tblLayout'))
                if existing is not None:
                    tblPr.remove(existing)
                tblPr.append(tblLayout)

                for i, t in enumerate(topics):
                    cell = table.rows[0].cells[i]
                    cell.text = t
                    if cell.text == 'nan':
                        cell.text = ''
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    for child in list(tcPr):
                        if child.tag == qn('w:textDirection'):
                            tcPr.remove(child)
                    rotation = parse_xml(r'<w:textDirection {} w:val="btLr"/>'.format(nsdecls('w')))
                    tcPr.append(rotation)
                    tcW = tcPr.find(qn('w:tcW'))
                    if tcW is None:
                        tcW = OxmlElement('w:tcW')
                        tcPr.append(tcW)
                    tcW.set(qn('w:w'), str(max_col_width_dxa))
                    tcW.set(qn('w:type'), 'dxa')

                for i, d in enumerate(dates):
                    table.rows[1].cells[i].text = d
                for i, g in enumerate(grades):
                    table.rows[2].cells[i].text = g

                max_len = max(len(str(t)) for t in topics) if topics else 1
                height_cm = max(max_len * HEIGHT_COEFF, BASE_HEIGHT_CM)
                row0 = table.rows[0]
                row0.height = Cm(height_cm)
                row0.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                for col_idx in range(ncols):
                    for r in table.rows:
                        cell = r.cells[col_idx]
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcW = tcPr.find(qn('w:tcW'))
                        if tcW is None:
                            tcW = OxmlElement('w:tcW')
                            tcPr.append(tcW)
                        tcW.set(qn('w:w'), str(max_col_width_dxa))
                        tcW.set(qn('w:type'), 'dxa')

            doc.add_page_break()

        doc.save(output_doc)

        st.success("Документ готов!")
        st.download_button(
            label="Скачать готовый документ",
            data=output_doc.getvalue(),
            file_name=f"Успеваемость от {datetime.datetime.today().strftime('%Y.%m.%d')} ({str(uploaded_file.name)}).docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
