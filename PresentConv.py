import re
import sys
from pathlib import Path
from lxml import etree as ET

from latex2mathml.converter import convert

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from markdown_parser import MarkdownParser


# Пространства имён XML
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"

ET.register_namespace("m", M)
ET.register_namespace("a", A)
ET.register_namespace("a14", A14)

def a14_element(name):
    return ET.Element(f"{{{A14}}}{name}")

def m_element(name):
    return ET.Element(f"{{{M}}}{name}")


def a_element(name):
    return ET.Element(f"{{{A}}}{name}")


class MathConverter:
    """Преобразование MathML в редактируемый Office Math."""

    def convert(self, latex):
        try:
            mathml = convert(latex)
            root = ET.fromstring(mathml)

            result = m_element("oMath")
            self.convert_node(root, result)

            return result

        except Exception as error:
            raise ValueError(
                f"Не удалось преобразовать формулу: {latex}"
            ) from error

    def convert_node(self, node, parent):
        tag = node.tag.split("}")[-1]

        children = list(node)

        if tag in ("math", "mrow", "semantics", "mstyle"):
            for child in children:
                if child.tag.split("}")[-1] == "annotation":
                    continue
                self.convert_node(child, parent)

            return

        if tag in ("mi", "mn", "mo", "mtext", "ms"):
            text = "".join(node.itertext())

            run = m_element("r")
            run_properties = m_element("rPr")

            if tag == "mi":
                run_properties.append(m_element("sty"))
                run_properties[-1].set(
                    f"{{{M}}}val", "i"
                )

            run.append(run_properties)

            text_element = m_element("t")
            text_element.text = text
            text_element.set(
                "{http://www.w3.org/XML/1998/namespace}space",
                "preserve"
            )

            run.append(text_element)
            parent.append(run)
            return

        if tag == "mfrac":
            if len(children) < 2:
                raise ValueError("Некорректная дробь MathML")

            fraction = m_element("f")
            fraction.append(m_element("fPr"))

            numerator = m_element("num")
            self.convert_node(children[0], numerator)
            fraction.append(numerator)

            denominator = m_element("den")
            self.convert_node(children[1], denominator)
            fraction.append(denominator)

            parent.append(fraction)
            return

        if tag == "msup":
            self.add_script(parent, children, "sSup")
            return

        if tag == "msub":
            self.add_script(parent, children, "sSub")
            return

        if tag == "msubsup":
            self.add_script(parent, children, "sSubSup")
            return

        if tag in ("msqrt", "mroot"):
            radical = m_element("rad")
            properties = m_element("radPr")

            if tag == "msqrt":
                hide = m_element("degHide")
                hide.set(f"{{{M}}}val", "1")
                properties.append(hide)

            radical.append(properties)

            degree = m_element("deg")

            if tag == "mroot" and len(children) > 1:
                self.convert_node(children[1], degree)

            radical.append(degree)

            expression = m_element("e")

            if children:
                self.convert_node(children[0], expression)

            radical.append(expression)
            parent.append(radical)
            return

        if tag == "mfenced":
            opening = node.attrib.get("open", "(")
            closing = node.attrib.get("close", ")")
            separator = node.attrib.get("separators", ",")

            self.add_text(parent, opening)

            for index, child in enumerate(children):
                if index > 0:
                    self.add_text(parent, separator)

                self.convert_node(child, parent)

            self.add_text(parent, closing)
            return

        if tag == "mover":
            self.add_accent(parent, children, "top")
            return

        if tag == "munder":
            self.add_script(parent, children, "sSub")
            return

        if tag == "munderover":
            self.add_script(parent, children, "sSubSup")
            return

        if tag in ("mspace", "none", "mprescripts"):
            return

        # Неизвестные элементы обрабатываем по содержимому.
        for child in children:
            self.convert_node(child, parent)

    def add_script(self, parent, children, kind):
        if len(children) < 2:
            raise ValueError("Некорректный индекс или степень")

        script = m_element(kind)
        script.append(m_element(kind + "Pr"))

        expression = m_element("e")
        self.convert_node(children[0], expression)
        script.append(expression)

        if kind in ("sSub", "sSubSup"):
            sub = m_element("sub")
            self.convert_node(children[1], sub)
            script.append(sub)

        if kind in ("sSup", "sSubSup"):
            sup = m_element("sup")
            self.convert_node(children[2] if kind == "sSubSup" else children[1], sup)
            script.append(sup)

        parent.append(script)

    def add_accent(self, parent, children, position):
        if len(children) < 2:
            return

        accent = m_element("acc")
        accent_properties = m_element("accPr")

        character = m_element("chr")
        character.set(
            f"{{{M}}}val",
            "\u0302" if position == "top" else "\u0332"
        )

        accent_properties.append(character)
        accent.append(accent_properties)

        expression = m_element("e")
        self.convert_node(children[0], expression)
        accent.append(expression)

        parent.append(accent)

    @staticmethod
    def add_text(parent, text):
        run = m_element("r")
        text_element = m_element("t")
        text_element.text = text
        run.append(text_element)
        parent.append(run)


class PowerPointCreator:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)

        self.parser = MarkdownParser()
        self.math = MathConverter()

        self.title_color = RGBColor(35, 75, 125)
        self.text_color = RGBColor(45, 45, 45)

    def create(self, markdown_file, output_file=None):
        slides = self.parser.parse(markdown_file)

        if output_file is None:
            output_file = Path(markdown_file).with_suffix(".pptx")

        # Сначала полностью формируем презентацию в памяти.
        # До успешного завершения всех операций файл не создаётся.
        for data in slides:
            self.add_slide(data)

        # Единственная операция сохранения.
        self.prs.save(str(output_file))

        print(f"Презентация создана: {output_file}")
        print(f"Количество слайдов: {len(slides)}")

    def add_paragraph_block(self, content):
        # Если текстового блока ещё нет, создаём его.
        if self.current_text_frame is None:
            self.current_text_frame = self.add_textbox(0.75)
            self.current_text_shape = self.slide.shapes[-1]
            self.current_text_count = 0
            paragraph = self.current_text_frame.paragraphs[0]
        else:
            # Новый абзац в уже существующем текстовом блоке.
            paragraph = self.current_text_frame.add_paragraph()

        self.current_text_count += 1

        paragraph.space_after = Pt(18)

        self.add_inline(paragraph, content)

        # Увеличиваем высоту текстового блока.
        self.current_text_shape.height = Inches(
            0.75 + (self.current_text_count - 1) * 0.45
        )

        self.y += 0.55

    def add_slide(self, data):
        self.current_text_frame = None
        self.current_text_shape = None
        self.current_text_count = 0

        slide = self.prs.slides.add_slide(
            self.prs.slide_layouts[6]
        )

        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(
            255, 255, 255
        )

        # Заголовок
        title_box = slide.shapes.add_textbox(
            Inches(0.65),
            Inches(0.3),
            Inches(12),
            Inches(0.8)
        )

        title = title_box.text_frame.paragraphs[0]
        title.text = data["title"]
        title.font.name = "Arial"
        title.font.size = Pt(30)
        title.font.bold = True
        title.font.color.rgb = self.title_color

        # Номер слайда
        number_box = slide.shapes.add_textbox(
            Inches(12),
            Inches(0.4),
            Inches(0.6),
            Inches(0.4)
        )

        number_paragraph = number_box.text_frame.paragraphs[0]
        number_paragraph.text = str(data["number"])
        number_paragraph.alignment = PP_ALIGN.RIGHT
        number_paragraph.font.size = Pt(16)
        number_paragraph.font.color.rgb = self.title_color

        # Линия под заголовком
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.65),
            Inches(1.2),
            Inches(12),
            Inches(0.025)
        )

        line.fill.solid()
        line.fill.fore_color.rgb = self.title_color
        line.line.fill.background()

        self.slide = slide
        self.y = 1.5

        for block in data["content"]:
            self.add_block(block)

    def add_textbox(self, height):
        box = self.slide.shapes.add_textbox(
            Inches(0.85),
            Inches(self.y),
            Inches(11.65),
            Inches(height)
        )

        frame = box.text_frame
        frame.word_wrap = True
        frame.margin_left = Inches(0.08)
        frame.margin_right = Inches(0.08)
        frame.margin_top = Inches(0.05)
        frame.margin_bottom = Inches(0.05)

        return frame

    def add_inline(self, paragraph, text, size=20):
        tokens = self.parser.md.parseInline(text)[0].children

        bold = False
        italic = False

        for token in tokens:
            if token.type == "strong_open":
                bold = True
                continue

            if token.type == "strong_close":
                bold = False
                continue

            if token.type == "em_open":
                italic = True
                continue

            if token.type == "em_close":
                italic = False
                continue

            if token.type == "s_open":
                continue

            if token.type == "s_close":
                continue

            if token.type == "code_inline":
                run = paragraph.add_run()
                run.text = token.content
                run.font.name = "Consolas"
                run.font.size = Pt(size)
                continue

            if token.type == "softbreak":
                run = paragraph.add_run()
                run.text = " "
                continue

            if token.type == "hardbreak":
                run = paragraph.add_run()
                run.text = "\n"
                continue

            if token.type != "text":
                continue

            # Распознаём LaTeX и Markdown-формулы.
            pattern = (
                r"(\\\(.*?\\\)|\\\[.*?\\\]|"
                r"(?<!\\)\$(?!\$).*?(?<!\\)\$)"
            )

            parts = re.split(pattern, token.content)

            for part in parts:
                if not part:
                    continue

                if part.startswith(r"\(") and part.endswith(r"\)"):
                    self.add_math_inline(
                        paragraph, part[2:-2]
                    )

                elif part.startswith("$") and part.endswith("$"):
                    self.add_math_inline(
                        paragraph, part[1:-1]
                    )

                elif part.startswith(r"\[") and part.endswith(r"\]"):
                    self.add_math_inline(
                        paragraph, part[2:-2]
                    )

                else:
                    run = paragraph.add_run()
                    run.text = part
                    run.font.name = "Arial"
                    run.font.size = Pt(size)
                    run.font.bold = bold
                    run.font.italic = italic
                    run.font.color.rgb = self.text_color

    def add_math_inline(self, paragraph, formula):
        math_element = self.math.convert(formula)

        math_shape = a14_element("m")
        math_shape.append(math_element)

        paragraph._p.append(math_shape)

    def add_math_block(self, formula):
        frame = self.add_textbox(0.8)
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER

        math_element = self.math.convert(formula)

        math_para = m_element("oMathPara")
        math_para.append(math_element)

        math_shape = a14_element("m")
        math_shape.append(math_para)

        paragraph._p.append(math_shape)

        self.y += 0.9

    def add_block(self, block):
        kind = block["type"]

        # При переходе к другому типу элемента
        # завершаем текущий текстовый блок.
        if kind != "paragraph":
            self.current_text_frame = None
            self.current_text_shape = None
            self.current_text_count = 0

        if kind == "paragraph":
            self.add_paragraph_block(block["content"])

        elif kind == "formula":
            self.add_math_block(block["content"])

        elif kind == "heading":
            frame = self.add_textbox(0.55)
            paragraph = frame.paragraphs[0]

            paragraph.text = block["content"]
            paragraph.font.name = "Arial"
            paragraph.font.size = Pt(
                25 if block["level"] == 2 else 22
            )
            paragraph.font.bold = True
            paragraph.font.color.rgb = self.title_color

            self.y += 0.6

        elif kind == "list":
            for index, item in enumerate(block["items"]):
                frame = self.add_textbox(0.55)
                paragraph = frame.paragraphs[0]

                prefix = (
                    f"{index + 1}. "
                    if block["ordered"]
                    else "• "
                )

                run = paragraph.add_run()
                run.text = prefix
                run.font.size = Pt(20)

                self.add_inline(paragraph, item)
                self.y += 0.55

        elif kind == "table":
            self.add_table(block["rows"])

        elif kind == "code":
            self.add_code(block["content"])

        elif kind == "quote":
            frame = self.add_textbox(0.8)
            paragraph = frame.paragraphs[0]

            self.add_inline(paragraph, block["content"])

            for run in paragraph.runs:
                run.font.italic = True
                run.font.color.rgb = RGBColor(100, 100, 100)

            self.y += 0.85

        elif kind == "separator":
            self.y += 0.2

    def add_table(self, rows):
        if not rows:
            return

        row_count = len(rows)
        col_count = max(len(row) for row in rows)

        table_height = min(0.55 * row_count, 3.5)

        table = self.slide.shapes.add_table(
            row_count,
            col_count,
            Inches(0.85),
            Inches(self.y),
            Inches(11.65),
            Inches(table_height)
        ).table

        for row_index, row in enumerate(rows):
            for col_index in range(col_count):
                cell = table.cell(row_index, col_index)

                cell.text = (
                    row[col_index]
                    if col_index < len(row)
                    else ""
                )

                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(15)

                        if row_index == 0:
                            run.font.bold = True
                            run.font.color.rgb = self.title_color

        self.y += table_height + 0.25

    def add_code(self, code):
        lines = code.splitlines()
        height = max(0.6, min(len(lines) * 0.25 + 0.2, 2.5))

        shape = self.slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.85),
            Inches(self.y),
            Inches(11.65),
            Inches(height)
        )

        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(240, 243, 247)
        shape.line.fill.background()

        frame = shape.text_frame
        frame.word_wrap = True
        frame.margin_left = Inches(0.15)
        frame.margin_top = Inches(0.1)

        paragraph = frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = code
        run.font.name = "Consolas"
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(45, 55, 65)

        self.y += height + 0.25


def main():
    if len(sys.argv) < 2:
        print(
            "Использование: "
            "python main.py input.md [output.pptx]"
        )
        sys.exit(1)

    input_file = Path(sys.argv[1])

    output_file = (
        Path(sys.argv[2])
        if len(sys.argv) > 2
        else input_file.with_suffix(".pptx")
    )

    if not input_file.is_file():
        print(f"Файл не найден: {input_file}")
        sys.exit(1)

    try:
        creator = PowerPointCreator()
        creator.create(input_file, output_file)

    except Exception as error:
        print(f"Ошибка конвертации: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
