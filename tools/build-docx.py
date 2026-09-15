#!/usr/bin/env python3
"""Сборка единого документа Word из ТЗ и приложений.

Использование:
    python3 tools/build-docx.py [--pandoc /path/to/pandoc] [--out FILE]

Требуется pandoc 3.x. Скрипт формирует временный reference.docx (альбомный A4,
сетка в таблицах, нумерация страниц) и склеивает три исходных файла в один .docx
с титульным листом и оглавлением.
"""

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TZ = ROOT / "docs" / "tz-checklists-contract-approval.md"
APPENDIX_A = ROOT / "docs" / "checklists-by-role.md"
APPENDIX_B = ROOT / "docs" / "checklists" / "contract-approval-checklists.yaml"
DEFAULT_OUT = ROOT / "docs" / "tz-checklists-contract-approval.docx"

PAGE_BREAK = '\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n'

# Ссылки на отдельные файлы в едином документе теряют смысл
LINK_REWRITES = {
    "`docs/checklists-by-role.md`": "Приложение А настоящего документа",
    "`docs/checklists/contract-approval-checklists.yaml`": "Приложение Б настоящего документа",
    "`docs/tz-checklists-contract-approval.md`": "основной текст настоящего ТЗ",
}

METADATA = """---
title: "Чек-листы согласования договоров в SimBASE"
subtitle: "Техническое задание. Версия 0.2 (проект). Включает Приложения А и Б"
date: "ТОО «Көркем Телеком», 2026"
lang: ru-RU
toc-title: "Содержание"
---
"""

APPENDIX_B_INTRO = """# Приложение Б. Машиночитаемый формат шаблонов чек-листов

Формат выгрузки и загрузки шаблонов чек-листов (основной текст ТЗ, FR-01 и FR-02).
Ниже приведено содержимое файла `contract-approval-checklists.yaml`: описание схемы,
четыре статуса пункта по п. 3.1 Правил, признаки карточки для условий применимости
и полностью описанные шаблоны трёх ролей как эталон формата.

"""


def rewrite_links(text: str) -> str:
    for src, dst in LINK_REWRITES.items():
        text = text.replace(src, dst)
    return text


def build_markdown() -> str:
    tz = rewrite_links(TZ.read_text(encoding="utf-8"))

    # Раздел со ссылками на отдельные файлы заменяем указанием на части документа
    tz = re.sub(
        r"## Приложения к ТЗ\n(?:.*\n)*?\Z",
        "## Приложения к ТЗ\n\n"
        "- **Приложение А.** Чек-листы согласования договора по ролям — приведено ниже.\n"
        "- **Приложение Б.** Машиночитаемый формат шаблонов чек-листов — приведено ниже.\n",
        tz,
    )

    appendix_a = rewrite_links(APPENDIX_A.read_text(encoding="utf-8"))
    appendix_b = APPENDIX_B_INTRO + "```yaml\n" + APPENDIX_B.read_text(encoding="utf-8") + "```\n"

    return METADATA + PAGE_BREAK + tz + PAGE_BREAK + appendix_a + PAGE_BREAK + appendix_b


def patch_reference(pandoc: str, work: Path) -> Path:
    """reference.docx pandoc по умолчанию: книжная ориентация, таблицы без сетки."""
    ref = work / "reference.docx"
    with ref.open("wb") as fh:
        subprocess.run([pandoc, "--print-default-data-file", "reference.docx"], stdout=fh, check=True)

    unpacked = work / "ref"
    with zipfile.ZipFile(ref) as z:
        z.extractall(unpacked)

    sect = (
        "<w:sectPr>"
        '<w:footerReference w:type="default" r:id="rIdFtr1"/>'
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
        '<w:pgMar w:top="851" w:right="851" w:bottom="851" w:left="851"'
        ' w:header="425" w:footer="425" w:gutter="0"/>'
        "</w:sectPr>"
    )
    document = unpacked / "word" / "document.xml"
    text = document.read_text(encoding="utf-8").replace("<w:sectPr />", sect, 1)
    document.write_text(text, encoding="utf-8")

    styles = unpacked / "word" / "styles.xml"
    text = styles.read_text(encoding="utf-8")
    text = text.replace(
        '<w:sz w:val="24" />\n        <w:szCs w:val="24" />',
        '<w:sz w:val="22" />\n        <w:szCs w:val="22" />',
        1,
    )
    text = text.replace(
        '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA" />',
        '<w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA" />',
        1,
    )
    borders = "".join(
        f'<w:{edge} w:val="single" w:sz="4" w:color="BFBFBF"/>'
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    text = text.replace(
        '<w:tblPr>\n      <w:tblInd w:w="0" w:type="dxa" />',
        f"<w:tblPr>\n      <w:tblBorders>{borders}</w:tblBorders>\n      "
        '<w:tblInd w:w="0" w:type="dxa" />',
        1,
    )
    text = text.replace(
        '<w:tcPr>\n          <w:vAlign w:val="bottom"/>',
        '<w:tcPr>\n          <w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>\n'
        '          <w:vAlign w:val="bottom"/>',
        1,
    )
    styles.write_text(text, encoding="utf-8")

    grey = '<w:rPr><w:sz w:val="18"/><w:color w:val="808080"/></w:rPr>'
    (unpacked / "word" / "footer1.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        f'  <w:p><w:pPr><w:jc w:val="center"/>{grey}</w:pPr>\n'
        f'    <w:r>{grey}<w:fldChar w:fldCharType="begin"/></w:r>\n'
        f'    <w:r>{grey}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>\n'
        f'    <w:r>{grey}<w:fldChar w:fldCharType="end"/></w:r>\n'
        "  </w:p>\n</w:ftr>\n",
        encoding="utf-8",
    )

    content_types = unpacked / "[Content_Types].xml"
    content_types.write_text(
        content_types.read_text(encoding="utf-8").replace(
            "</Types>",
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxml'
            'formats-officedocument.wordprocessingml.footer+xml" /></Types>',
        ),
        encoding="utf-8",
    )

    rels = unpacked / "word" / "_rels" / "document.xml.rels"
    rels.write_text(
        rels.read_text(encoding="utf-8").replace(
            "</Relationships>",
            '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/footer" Id="rIdFtr1" Target="footer1.xml" /></Relationships>',
        ),
        encoding="utf-8",
    )

    patched = work / "reference-landscape.docx"
    with zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(unpacked.rglob("*")):
            if path.is_file():
                z.write(path, path.relative_to(unpacked).as_posix())
    return patched


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pandoc", default=shutil.which("pandoc") or "pandoc")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args()

    if not shutil.which(args.pandoc) and not Path(args.pandoc).is_file():
        print(f"pandoc не найден: {args.pandoc}", file=sys.stderr)
        return 1

    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        combined = work / "combined.md"
        combined.write_text(build_markdown(), encoding="utf-8")
        reference = patch_reference(args.pandoc, work)
        subprocess.run(
            [
                args.pandoc,
                str(combined),
                "--from=markdown",
                "--to=docx",
                f"--reference-doc={reference}",
                "--toc",
                "--toc-depth=2",
                f"--output={args.out}",
            ],
            check=True,
        )

    print(f"готово: {args.out} ({args.out.stat().st_size // 1024} КБ)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
