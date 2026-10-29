"""Strip semantic styles/direct typography from the complete synthetic thesis.

Only standard-library modules are used. This is test material, never a student paper.
Page/section boundaries, tables, equations, pictures, text and other ZIP parts remain.
"""
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
HERE = Path(__file__).resolve().parent


def remove(node) -> None:
    node.parentNode.removeChild(node)


def plain_run_properties(document):
    properties = document.createElementNS(W, "w:rPr")
    fonts = document.createElementNS(W, "w:rFonts")
    for name, value in (("eastAsia", "黑体"), ("ascii", "Arial"), ("hAnsi", "Arial")):
        fonts.setAttributeNS(W, "w:" + name, value)
    properties.appendChild(fonts)
    size = document.createElementNS(W, "w:sz")
    size.setAttributeNS(W, "w:val", "22")
    properties.appendChild(size)
    return properties


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    arguments = parser.parse_args()
    source = HERE / "synthetic_complete_thesis.docx"
    output = arguments.output.resolve()
    if output == source.resolve() or output.exists() or output.suffix.lower() != ".docx":
        parser.error("Choose a new .docx output; source and existing files are never overwritten.")
    if not output.parent.is_dir():
        parser.error("Output parent directory must exist.")
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    with ZipFile(source) as package:
        original = package.read("word/document.xml")
        document = minidom.parseString(original)
        before_text = [node.firstChild.data if node.firstChild else "" for node in document.getElementsByTagNameNS(W, "t")]
        for properties in list(document.getElementsByTagNameNS(W, "pPr")):
            for child in list(properties.childNodes):
                if child.nodeType == child.ELEMENT_NODE and child.localName not in ("sectPr", "pageBreakBefore"):
                    properties.removeChild(child)
        for properties in list(document.getElementsByTagNameNS(W, "rPr")):
            remove(properties)
        styles = minidom.parseString(package.read("word/styles.xml"))
        for defaults in list(styles.getElementsByTagNameNS(W, "docDefaults")):
            remove(defaults)
        defaults = styles.createElementNS(W, "w:docDefaults")
        run_defaults = styles.createElementNS(W, "w:rPrDefault")
        run_defaults.appendChild(plain_run_properties(styles))
        defaults.appendChild(run_defaults)
        styles.documentElement.insertBefore(defaults, styles.documentElement.firstChild)
        for style in styles.getElementsByTagNameNS(W, "style"):
            if style.getAttributeNS(W, "type") == "paragraph" and style.getAttributeNS(W, "default") in ("1", "true", "on"):
                for child in list(style.childNodes):
                    if child.nodeType == child.ELEMENT_NODE and child.localName in ("rPr", "pPr"):
                        style.removeChild(child)
                style.appendChild(plain_run_properties(styles))
        after_text = [node.firstChild.data if node.firstChild else "" for node in document.getElementsByTagNameNS(W, "t")]
        if before_text != after_text:
            raise RuntimeError("Test-fixture transform changed text.")
        with ZipFile(output, "x") as result:
            for entry in package.infolist():
                content = document.toxml(encoding="utf-8") if entry.filename == "word/document.xml" else (
                    styles.toxml(encoding="utf-8") if entry.filename == "word/styles.xml" else package.read(entry.filename))
                result.writestr(entry, content)
    if source_hash != hashlib.sha256(source.read_bytes()).hexdigest():
        raise RuntimeError("Original fixture changed.")
    print(f"Synthetic unstyled thesis: {output}")
    print(f"Paragraphs: {len(document.getElementsByTagNameNS(W, 'p'))}; text nodes preserved: {len(before_text)}")
    print(f"SHA256: {hashlib.sha256(output.read_bytes()).hexdigest()}")


if __name__ == "__main__":
    main()
