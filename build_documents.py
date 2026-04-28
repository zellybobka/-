from __future__ import annotations

import textwrap
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


BASE_DIR = Path(__file__).resolve().parent
PROGRAM_NAME = "Базовая СФМ ВОЛС"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)

    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(14)


def add_cover_page(document: Document, title: str, subtitle: str) -> None:
    blocks = [
        "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
        "Федеральное государственное автономное образовательное учреждение высшего образования",
        "«Санкт-Петербургский государственный университет аэрокосмического приборостроения»",
        "КАФЕДРА 33",
    ]
    for index, text in enumerate(blocks):
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if index == 0:
            p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.bold = True if index in {0, 3} else False

    for _ in range(4):
        document.add_paragraph()

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(subtitle)
    run.bold = True

    for _ in range(6):
        document.add_paragraph()

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run("Студент: ____________________\n")
    p.add_run("Группа: ____________________\n")
    p.add_run("Преподаватель: ____________________")

    for _ in range(6):
        document.add_paragraph()

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Санкт-Петербург\n2026")

    document.add_page_break()


def add_heading(document: Document, text: str) -> None:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    run.bold = True


def add_body_paragraph(document: Document, text: str) -> None:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(1.25)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    p.add_run(text)


def build_abstract() -> Path:
    document = Document()
    configure_document(document)
    add_cover_page(document, "РЕФЕРАТ", PROGRAM_NAME)
    add_heading(document, "Реферат")

    paragraphs = [
        "Программа для ЭВМ «Базовая СФМ ВОЛС» предназначена для составления базовой структурно-функциональной модели волоконно-оптической линии связи в табличной форме в соответствии с методическими рекомендациями по индивидуальному заданию.",
        "Приложение реализовано как кроссплатформенная настольная программа на языке Python с графическим интерфейсом Tkinter и может использоваться в операционных системах Windows и Linux при наличии Python 3.10 и выше.",
        "Функциональные возможности программы включают ввод названия объекта моделирования, добавление и редактирование строк таблицы, задание функций технических систем, требований, ограничений и диапазонов показателей качества Qmin и Qmax, а также сохранение модели в формате JSON и экспорт итоговой таблицы в CSV.",
        "Практическая ценность разработки заключается в ускорении подготовки базовой СФМ ВОЛС, унификации структуры данных и снижении числа ошибок при оформлении таблиц, используемых для дальнейшего анализа качества линии связи.",
    ]
    for text in paragraphs:
        add_body_paragraph(document, text)

    path = BASE_DIR / "Реферат_Базовая_СФМ_ВОЛС.docx"
    document.save(path)
    return path


def build_user_guide() -> Path:
    document = Document()
    configure_document(document)
    add_cover_page(document, "РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ", PROGRAM_NAME)
    add_heading(document, "Руководство пользователя")

    intro = [
        "Программа «Базовая СФМ ВОЛС» предназначена для подготовки базовой структурно-функциональной модели волоконно-оптической линии связи в виде таблицы. Интерфейс состоит из области ввода параметров строки и области просмотра сформированной таблицы.",
        "Для начала работы необходимо установить Python 3.10 или более новой версии. В Windows следует запустить файл install_windows.bat, затем run_windows.bat. В Linux нужно выдать права на выполнение файлам install_linux.sh и run_linux.sh, после чего последовательно выполнить их в терминале.",
    ]
    for text in intro:
        add_body_paragraph(document, text)

    steps = [
        "1. После запуска программы в поле «Название объекта» введите наименование моделируемой ВОЛС.",
        "2. Заполните поля строки: техническая система, функция, требование или ограничение, показатель качества, значения Qmin и Qmax, признак критичности параметра.",
        "3. Нажмите кнопку «Добавить строку». Строка появится в итоговой таблице справа.",
        "4. Для изменения записи выделите строку в таблице, отредактируйте данные в форме и нажмите «Обновить строку».",
        "5. Для удаления строки выделите её и нажмите «Удалить строку».",
        "6. Кнопка «Шаблон по умолчанию» загружает пример базовой модели ВОЛС.",
        "7. Для сохранения модели в рабочем формате используйте кнопку «Сохранить JSON». Для переноса таблицы в отчёт или электронные таблицы используйте кнопку «Экспорт CSV».",
    ]
    for text in steps:
        add_body_paragraph(document, text)

    add_body_paragraph(
        document,
        "При вводе данных программа проверяет корректность значений Qmin и Qmax. Если значение Qmin больше Qmax либо введены нечисловые данные, пользователю выводится предупреждение."
    )

    path = BASE_DIR / "Руководство_пользователя_Базовая_СФМ_ВОЛС.docx"
    document.save(path)
    return path


def build_code_doc() -> Path:
    document = Document()
    configure_document(document)
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    add_cover_page(document, "ПРОГРАММНЫЙ КОД", PROGRAM_NAME)
    add_heading(document, "Листинг исходного кода")

    source_text = (BASE_DIR / "vols_sfm_app.py").read_text(encoding="utf-8")
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(source_text)
    run.font.name = "Courier New"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Courier New")
    run.font.size = Pt(9)

    path = BASE_DIR / "Программный_код_Базовая_СФМ_ВОЛС.docx"
    document.save(path)
    return path


def build_distribution_zip() -> Path:
    archive_path = BASE_DIR / "Базовая_СФМ_ВОЛС_программа.zip"
    files = [
        "README.md",
        "requirements.txt",
        "sample_model.json",
        "vols_sfm_app.py",
        "install_windows.bat",
        "run_windows.bat",
        "install_linux.sh",
        "run_linux.sh",
    ]
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name in files:
            archive.write(BASE_DIR / name, arcname=name)
    return archive_path


def main() -> None:
    build_abstract()
    build_user_guide()
    build_code_doc()
    build_distribution_zip()


if __name__ == "__main__":
    main()
