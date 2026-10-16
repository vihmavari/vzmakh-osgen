import datetime
import hashlib
from io import BytesIO

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm


# -----------------------------------------------------------------------------
# Интерфейс
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Генератор отчетов", page_icon="🖨", layout="wide")
st.title("Генератор отчетов по успеваемости")
st.caption(
    "Загрузите журнал, настройте список предметов и при необходимости объедините "
    "варианты имён учеников. Затем создайте персональные табели в Word."
)


# -----------------------------------------------------------------------------
# Боковая панель: параметры отчёта
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Настройки отчета")

st.sidebar.subheader("Фильтр данных")
start_date = st.sidebar.date_input(
    "Показывать оценки начиная с:",
    value=datetime.date(2026, 1, 1),
    help="Оценки за даты ранее выбранной не попадут в обратную связь.",
)

use_only_surname = st.sidebar.toggle(
    "Использовать только фамилии",
    value=False,
    help=(
        "Отчёт будет группироваться только по первому слову из ФИО. "
        "Используйте, только если в классе нет однофамильцев."
    ),
)

st.sidebar.subheader("Внешний вид")
rotate_text = st.sidebar.toggle("Поворачивать текст тем вертикально", value=True)
compact_mode = st.sidebar.toggle(
    "Компактный режим (Beta)",
    value=False,
    help="Размещает таблицы в ряд для экономии бумаги.",
)
orientation = st.sidebar.selectbox("Ориентация страницы", ["Книжная", "Альбомная"])
show_dates = st.sidebar.toggle("Включить строку дат", value=True)
separator_type = st.sidebar.selectbox(
    "Разделитель между учениками",
    ["Разрыв страницы", "Пустые строки (параграфы)"],
)

num_paragraphs = 1
if separator_type == "Пустые строки (параграфы)":
    num_paragraphs = st.sidebar.number_input(
        "Количество пустых строк", min_value=1, max_value=10, value=2
    )


# -----------------------------------------------------------------------------
# Константы таблицы
# -----------------------------------------------------------------------------
MAX_WIDTH_CM = 2
BASE_HEIGHT_CM = 1.5
HEIGHT_COEFF = 0.07
MAX_TOPIC_HEIGHT_CM = 5.5

if orientation == "Альбомная":
    PAGE_WIDTH_CM = 27.5  # ширина листа минус поля
else:
    PAGE_WIDTH_CM = 19.0


def cm_to_dxa(cm):
    """Перевод сантиметров в единицы измерения Word (twips/dxa)."""
    inches = cm / 2.54
    points = inches * 72
    return int(round(points * 20))


def format_grade(val):
    """Форматирование оценки; сохраняет исходный формат дат в ячейках оценок."""
    if pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%d/%m").lstrip("0").replace("/0", "/")
    return str(val).strip()


max_col_width_dxa = cm_to_dxa(MAX_WIDTH_CM)


# -----------------------------------------------------------------------------
# Вспомогательные функции чтения исходного Excel
# -----------------------------------------------------------------------------
def read_student_identifiers(workbook_bytes, subject_sheets, surname_only):
    """Собирает уникальные имена учеников из строк журналов, начиная со строки 6."""
    identifiers = set()

    for sheet in subject_sheets:
        try:
            df = pd.read_excel(BytesIO(workbook_bytes), sheet_name=sheet, header=None)
            students = df.iloc[5:, :]

            for _, row in students.iterrows():
                if len(row) <= 1:
                    continue

                value = row.iloc[1]
                if pd.isna(value):
                    continue

                raw_name = str(value).strip()
                if not raw_name or raw_name.lower() == "nan":
                    continue

                identifier = raw_name.split()[0] if surname_only else raw_name
                if identifier:
                    identifiers.add(identifier)
        except Exception:
            # Поведение соответствует исходной версии: листы, которые не удаётся
            # прочитать как таблицу оценок, пропускаются.
            continue

    return sorted(identifiers, key=str.casefold)


def get_subject_sheets(workbook_bytes):
    """Берёт перечень предметов из первого (навигационного) листа."""
    xls = pd.ExcelFile(BytesIO(workbook_bytes))
    sheet_names = xls.sheet_names
    if not sheet_names:
        return [], []

    index_df = pd.read_excel(BytesIO(workbook_bytes), sheet_name=sheet_names[0], header=None)
    listed_subjects = index_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()

    # Сохраняем порядок исходного перечня, убирая пустые, повторные и отсутствующие листы.
    subject_sheets = []
    seen = set()
    for subject in listed_subjects:
        if subject and subject in sheet_names and subject not in seen:
            subject_sheets.append(subject)
            seen.add(subject)

    return sheet_names, subject_sheets


# -----------------------------------------------------------------------------
# Загрузка Excel и две вкладки постобработки
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Выберите Excel файл (журнал класса)", type=["xlsx"]
)

if uploaded_file:
    workbook_bytes = uploaded_file.getvalue()
    file_signature = hashlib.sha1(workbook_bytes).hexdigest()[:12]

    try:
        sheet_names, all_subject_sheets = get_subject_sheets(workbook_bytes)

        if not all_subject_sheets:
            st.error(
                "В первом листе Excel не найден перечень предметов, соответствующих "
                "листам этой книги. Проверьте навигационный лист журнала."
            )
        else:
            st.subheader("Настройка данных")
            subject_tab, student_tab = st.tabs(["📚 1. Предметы", "👥 2. Ученики"])

            with subject_tab:
                st.markdown(
                    "Выберите предметы, которые должны попасть в обратную связь. "
                    "Удалённые из выбора предметы будут полностью пропущены при генерации."
                )
                selected_subjects = st.multiselect(
                    "Предметы в отчёте",
                    options=all_subject_sheets,
                    default=all_subject_sheets,
                    key=f"subjects_{file_signature}",
                    help="Уберите из списка предметы, которые не нужно включать в табели.",
                )
                st.caption(
                    f"Выбрано предметов: {len(selected_subjects)} из {len(all_subject_sheets)}."
                )

            # Список идентификаторов читается из всех предметов, а не только выбранных:
            # так таблица сопоставления имён не исчезает при переключении предметов.
            student_identifiers = read_student_identifiers(
                workbook_bytes, all_subject_sheets, use_only_surname
            )

            with student_tab:
                st.markdown(
                    "Если один ученик записан в разных предметах по-разному, укажите "
                    "для этих строк одинаковое **«Имя в отчёте»**. В Word они будут "
                    "собраны в один табель. Здесь также можно исправить имя для печати."
                )

                if student_identifiers:
                    mapping_df = pd.DataFrame(
                        {
                            "Имя из журнала": student_identifiers,
                            "Имя в отчёте": student_identifiers,
                        }
                    )
                    edited_mapping = st.data_editor(
                        mapping_df,
                        key=f"student_mapping_{file_signature}_{int(use_only_surname)}",
                        hide_index=True,
                        num_rows="fixed",
                        use_container_width=True,
                        column_config={
                            "Имя из журнала": st.column_config.TextColumn(
                                "Имя из журнала", disabled=True
                            ),
                            "Имя в отчёте": st.column_config.TextColumn(
                                "Имя в отчёте",
                                help=(
                                    "Одинаковое итоговое имя объединяет записи в один табель. "
                                    "Не оставляйте поле пустым."
                                ),
                            ),
                        },
                    )
                else:
                    st.info("Не удалось найти имена учеников в строках журналов.")
                    edited_mapping = pd.DataFrame(
                        columns=["Имя из журнала", "Имя в отчёте"]
                    )

            # Строим словарь исходное имя -> итоговое имя. Если поле случайно очищено,
            # используем исходный идентификатор, чтобы ученик не потерялся.
            student_name_map = {}
            for _, mapping_row in edited_mapping.iterrows():
                source_name = str(mapping_row.get("Имя из журнала", "")).strip()
                target_value = mapping_row.get("Имя в отчёте", source_name)
                target_name = (
                    ""
                    if pd.isna(target_value)
                    else str(target_value).strip()
                )
                student_name_map[source_name] = target_name or source_name

            if not selected_subjects:
                st.warning("Выберите хотя бы один предмет на вкладке «Предметы».")

            if st.button(
                "🚀 Создать ОС",
                type="primary",
                disabled=not selected_subjects,
            ):
                progress_bar = st.progress(0)
                status_text = st.empty()
                output_doc = BytesIO()
                results = []

                try:
                    # ------------------ ЭТАП 1: СБОР ДАННЫХ ------------------
                    for i, sheet in enumerate(selected_subjects):
                        progress_bar.progress(i / len(selected_subjects))
                        status_text.text(f"🔍 Обработка: {sheet}")

                        if sheet not in sheet_names:
                            continue

                        try:
                            df = pd.read_excel(
                                BytesIO(workbook_bytes), sheet_name=sheet, header=None
                            )
                            topics = df.iloc[0, 3:]
                            dates_raw = df.iloc[1, 3:]
                            students = df.iloc[5:, :]

                            for _, row in students.iterrows():
                                if len(row) <= 1 or pd.isna(row.iloc[1]):
                                    continue

                                raw_name = str(row.iloc[1]).strip()
                                if not raw_name or raw_name.lower() == "nan":
                                    continue

                                source_name = (
                                    raw_name.split()[0] if use_only_surname else raw_name
                                )
                                student_name = student_name_map.get(
                                    source_name, source_name
                                )
                                if not student_name:
                                    student_name = source_name

                                for topic, date_raw, grade in zip(
                                    topics, dates_raw, row.iloc[3:]
                                ):
                                    try:
                                        current_date = pd.to_datetime(date_raw).date()
                                        if current_date < start_date:
                                            continue

                                        formatted_grade = format_grade(grade)
                                        if (
                                            not formatted_grade
                                            or formatted_grade.lower() == "nan"
                                        ):
                                            continue

                                        results.append(
                                            {
                                                "ФИО": student_name,
                                                "Предмет": str(sheet).strip(),
                                                "Тема": str(topic).strip(),
                                                "Дата": current_date,
                                                "Оценка": formatted_grade,
                                            }
                                        )
                                    except Exception:
                                        # Некорректная дата или пустая ячейка — пропускаем оценку.
                                        continue
                        except Exception:
                            # Пропускаем лист, который не удалось разобрать.
                            continue

                    progress_bar.progress(1.0)

                    if not results:
                        status_text.empty()
                        st.error(
                            "Оценок не найдено. Проверьте дату начала, выбранные предметы "
                            "и исходный журнал."
                        )
                    else:
                        with st.spinner("Формирование документа..."):
                            doc = Document()

                            # Настройка ориентации и полей страницы.
                            section = doc.sections[0]
                            if orientation == "Альбомная":
                                section.orientation = WD_ORIENT.LANDSCAPE
                                section.page_width, section.page_height = Cm(29.7), Cm(21.0)

                            margin = 0.8 if compact_mode else 1.0
                            section.top_margin = section.bottom_margin = Cm(margin)
                            section.left_margin = section.right_margin = Cm(margin)

                            full = pd.DataFrame(results).sort_values(
                                ["ФИО", "Предмет", "Дата"]
                            )
                            student_groups = list(full.groupby("ФИО", sort=True))

                            # ------------------ ЭТАП 2: ФОРМИРОВАНИЕ WORD ------------------
                            for idx, (student, df_student) in enumerate(student_groups):
                                status_text.text(f"📄 Оформление: {student}")
                                doc.add_heading(student, level=2)

                                subjects = []
                                for subject_name, df_subject in df_student.groupby(
                                    "Предмет", sort=True
                                ):
                                    topic_list = df_subject["Тема"].tolist()
                                    formatted_dates = [
                                        date_value.strftime("%d.%m")
                                        for date_value in df_subject["Дата"]
                                    ]
                                    subjects.append(
                                        {
                                            "name": subject_name,
                                            "topics": topic_list,
                                            "dates": formatted_dates,
                                            "grades": df_subject["Оценка"].tolist(),
                                            "w": len(topic_list) * MAX_WIDTH_CM,
                                        }
                                    )

                                if compact_mode:
                                    # -------- Компактный режим: таблицы предметов в ряд --------
                                    rows, current_row, current_width = [], [], 0
                                    for subject in subjects:
                                        if (
                                            current_width + subject["w"] > PAGE_WIDTH_CM
                                            and current_row
                                        ):
                                            rows.append(current_row)
                                            current_row, current_width = [], 0
                                        current_row.append(subject)
                                        current_width += subject["w"] + 0.5
                                    if current_row:
                                        rows.append(current_row)

                                    for row_subjects in rows:
                                        total_cols = sum(
                                            len(subject["topics"])
                                            for subject in row_subjects
                                        )
                                        container = doc.add_table(
                                            rows=1, cols=len(row_subjects)
                                        )
                                        container.autofit = False

                                        for col_idx, subject in enumerate(row_subjects):
                                            cell = container.rows[0].cells[col_idx]
                                            cell_width = int(
                                                (
                                                    len(subject["topics"])
                                                    / total_cols
                                                )
                                                * cm_to_dxa(PAGE_WIDTH_CM)
                                            )
                                            cell._tc.get_or_add_tcPr().get_or_add_tcW().set(
                                                qn("w:w"), str(cell_width)
                                            )

                                            cell.paragraphs[0].add_run(subject["name"]).bold = True
                                            row_count = 3 if show_dates else 2
                                            inner = cell.add_table(
                                                rows=row_count,
                                                cols=len(subject["topics"]),
                                            )
                                            inner.style = "Table Grid"
                                            inner._tbl.tblPr.append(
                                                parse_xml(
                                                    r'<w:tblLayout {} w:type="fixed"/>'.format(
                                                        nsdecls("w")
                                                    )
                                                )
                                            )

                                            for col_idx_inner, (topic, date_text, grade_text) in enumerate(
                                                zip(
                                                    subject["topics"],
                                                    subject["dates"],
                                                    subject["grades"],
                                                )
                                            ):
                                                inner.rows[0].cells[col_idx_inner].text = str(topic)
                                                if rotate_text:
                                                    tc_pr = inner.rows[0].cells[
                                                        col_idx_inner
                                                    ]._tc.get_or_add_tcPr()
                                                    tc_pr.append(
                                                        parse_xml(
                                                            r'<w:textDirection {} w:val="btLr"/>'.format(
                                                                nsdecls("w")
                                                            )
                                                        )
                                                    )

                                                if show_dates:
                                                    inner.rows[1].cells[col_idx_inner].text = str(
                                                        date_text
                                                    )
                                                inner.rows[-1].cells[col_idx_inner].text = str(
                                                    grade_text
                                                )

                                                for row_idx in range(row_count):
                                                    cell_width_element = (
                                                        inner.rows[row_idx]
                                                        .cells[col_idx_inner]
                                                        ._tc.get_or_add_tcPr()
                                                        .get_or_add_tcW()
                                                    )
                                                    cell_width_element.set(
                                                        qn("w:w"), str(max_col_width_dxa)
                                                    )

                                            if rotate_text and subject["topics"]:
                                                calculated_height = (
                                                    max(len(str(topic)) for topic in subject["topics"])
                                                    * HEIGHT_COEFF
                                                )
                                                final_height = max(
                                                    BASE_HEIGHT_CM,
                                                    min(calculated_height, MAX_TOPIC_HEIGHT_CM),
                                                )
                                                inner.rows[0].height = Cm(final_height)
                                                inner.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                                    doc.add_paragraph()

                                else:
                                    # -------- Классический режим: каждый предмет отдельно --------
                                    for subject in subjects:
                                        doc.add_heading(subject["name"], level=3)
                                        row_count = 3 if show_dates else 2
                                        table = doc.add_table(
                                            rows=row_count,
                                            cols=len(subject["topics"]),
                                        )
                                        table.style = "Table Grid"

                                        for col_idx, (topic, date_text, grade_text) in enumerate(
                                            zip(
                                                subject["topics"],
                                                subject["dates"],
                                                subject["grades"],
                                            )
                                        ):
                                            table.rows[0].cells[col_idx].text = str(topic)
                                            if rotate_text:
                                                tc_pr = table.rows[0].cells[
                                                    col_idx
                                                ]._tc.get_or_add_tcPr()
                                                tc_pr.append(
                                                    parse_xml(
                                                        r'<w:textDirection {} w:val="btLr"/>'.format(
                                                            nsdecls("w")
                                                        )
                                                    )
                                                )

                                            if show_dates:
                                                table.rows[1].cells[col_idx].text = str(date_text)
                                            table.rows[-1].cells[col_idx].text = str(grade_text)

                                            for row_idx in range(row_count):
                                                table.rows[row_idx].cells[
                                                    col_idx
                                                ]._tc.get_or_add_tcPr().get_or_add_tcW().set(
                                                    qn("w:w"), str(max_col_width_dxa)
                                                )

                                        if rotate_text and subject["topics"]:
                                            table.rows[0].height = Cm(
                                                max(
                                                    max(
                                                        len(str(topic))
                                                        for topic in subject["topics"]
                                                    )
                                                    * HEIGHT_COEFF,
                                                    BASE_HEIGHT_CM,
                                                )
                                            )

                                if idx < len(student_groups) - 1:
                                    if separator_type == "Разрыв страницы":
                                        doc.add_page_break()
                                    else:
                                        for _ in range(int(num_paragraphs)):
                                            doc.add_paragraph()

                            doc.save(output_doc)

                        progress_bar.progress(1.0)
                        status_text.empty()
                        st.success(
                            f"✅ Готово! Ученики: {len(student_groups)}; "
                            f"предметы в выборке: {len(selected_subjects)}."
                        )

                        filename = (
                            f"Обратная_Связь_{datetime.datetime.now().strftime('%d.%m')} "
                            f"({uploaded_file.name.rsplit('.', 1)[0]}).docx"
                        )
                        st.download_button(
                            label="📥 Скачать обратную связь",
                            data=output_doc.getvalue(),
                            file_name=filename,
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "wordprocessingml.document"
                            ),
                        )

                except Exception as exc:
                    status_text.empty()
                    st.error(f"Критическая ошибка: {exc}")
                    st.exception(exc)

    except Exception as exc:
        st.error(f"Не удалось прочитать Excel-файл: {exc}")
        st.exception(exc)
