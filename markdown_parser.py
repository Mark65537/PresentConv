
import re

from markdown_it import MarkdownIt


class MarkdownParser:
    def __init__(self):
        self.md = MarkdownIt("default").enable([
            "table",
            "strikethrough"
        ])

    def parse(self, filename):
        with open(filename, "r", encoding="utf-8-sig") as file:
            text = file.read()

        sections = self.split_slides(text)

        if not sections:
            raise ValueError("В документе не найдены слайды.")

        slides = []

        for number, title, content in sections:
            slides.append({
                "number": number,
                "title": title,
                "content": self.parse_content(content)
            })

        return slides

    def split_slides(self, text):
        sections = []

        current = None
        content = []
        in_code = False

        for line in text.splitlines():
            if line.strip().startswith("```"):
                in_code = not in_code

            match = None

            if not in_code:
                match = re.match(
                    r"^#{1,6}\s+"
                    r"(?:слайд\s+)?"
                    r"(\d+)[.)]\s*(.+?)\s*$",
                    line,
                    re.IGNORECASE
                )

            if match:
                if current is not None:
                    sections.append((
                        current[0],
                        current[1],
                        "\n".join(content)
                    ))

                current = (
                    int(match.group(1)),
                    match.group(2).strip()
                )

                content = []

            else:
                if current is None:
                    if line.strip():
                        raise ValueError(
                            "Обнаружен текст до первого "
                            "пронумерованного заголовка: "
                            + line
                        )

                else:
                    content.append(line)

        if current is not None:
            sections.append((
                current[0],
                current[1],
                "\n".join(content)
            ))

        return sections

    def parse_content(self, text):
        # Извлекаем блочные формулы до разбора Markdown.
        # Это позволяет сохранить их содержимое без изменений.
        formulas = {}

        def replace_formula(match):
            key = f"MATHBLOCKTOKEN{len(formulas)}"
            formulas[key] = match.group(1).strip()
            return f"\n{key}\n"

        text = re.sub(
            r"(?m)^[ \t]*\$\$\s*\n(.*?)\n[ \t]*\$\$[ \t]*$",
            replace_formula,
            text,
            flags=re.DOTALL
        )

        text = re.sub(
            r"(?m)^[ \t]*\\\[\s*\n(.*?)\n[ \t]*\\\][ \t]*$",
            replace_formula,
            text,
            flags=re.DOTALL
        )

        # Однострочные блочные формулы.
        text = re.sub(
            r"(?m)^[ \t]*\$\$(.+?)\$\$[ \t]*$",
            replace_formula,
            text
        )

        text = re.sub(
            r"(?m)^[ \t]*\\\[(.+?)\\\][ \t]*$",
            replace_formula,
            text
        )

        tokens = self.md.parse(text)
        blocks = []
        i = 0

        while i < len(tokens):
            token = tokens[i]

            if token.type in ("fence", "code_block"):
                blocks.append({
                    "type": "code",
                    "content": token.content,
                    "language": token.info.strip()
                })
                i += 1
                continue

            if token.type == "paragraph_open":
                inline = tokens[i + 1].content.strip()

                if inline in formulas:
                    blocks.append({
                        "type": "formula",
                        "content": formulas[inline]
                    })
                else:
                    blocks.append({
                        "type": "paragraph",
                        "content": inline
                    })

                i += 3
                continue

            if token.type == "heading_open":
                inline = tokens[i + 1].content.strip()

                blocks.append({
                    "type": "heading",
                    "level": int(token.tag[1]),
                    "content": inline
                })

                i += 3
                continue

            if token.type in ("bullet_list_open", "ordered_list_open"):
                ordered = token.type == "ordered_list_open"
                items = []

                i += 1

                while i < len(tokens) and tokens[i].type != (
                    "ordered_list_close" if ordered
                    else "bullet_list_close"
                ):
                    if tokens[i].type == "inline":
                        items.append(tokens[i].content)

                    i += 1

                blocks.append({
                    "type": "list",
                    "ordered": ordered,
                    "items": items
                })

                i += 1
                continue

            if token.type == "table_open":
                rows = []
                current_row = []
                i += 1

                while i < len(tokens) and tokens[i].type != "table_close":
                    if tokens[i].type == "tr_open":
                        current_row = []

                    elif tokens[i].type == "inline":
                        current_row.append(tokens[i].content)

                    elif tokens[i].type == "tr_close":
                        rows.append(current_row)

                    i += 1

                blocks.append({
                    "type": "table",
                    "rows": rows
                })

                i += 1
                continue

            if token.type == "blockquote_open":
                quote = []
                i += 1

                while i < len(tokens) and tokens[i].type != "blockquote_close":
                    if tokens[i].type == "inline":
                        quote.append(tokens[i].content)
                    i += 1

                blocks.append({
                    "type": "quote",
                    "content": "\n".join(quote)
                })

                i += 1
                continue

            if token.type == "hr":
                blocks.append({"type": "separator"})

            i += 1

        return blocks
