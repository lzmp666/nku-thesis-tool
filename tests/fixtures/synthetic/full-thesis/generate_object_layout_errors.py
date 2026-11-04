"""Derive a complete synthetic thesis with intentionally oversized figure/table.

Only generated fixtures are accepted. No student or user document is read.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import argparse
from lxml import etree as ET

HERE = Path(__file__).resolve().parent
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}


def make_fixture(output: Path):
    source = HERE / "synthetic_complete_thesis.docx"
    if output.resolve() == source.resolve() or output.exists():
        raise ValueError("Output must be a new synthetic fixture, never the source.")
    with ZipFile(source) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        body = root.find("w:body", NS)
        # The older fixture generator appended custom borders after vAlign and
        # omitted the required shading pattern. Normalize only this derived fixture.
        tc_order = {name: index for index, name in enumerate([
            "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
            "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"])}
        for cell_props in body.findall(".//w:tcPr", NS):
            for shade in cell_props.findall("w:shd", NS):
                if shade.get("{" + NS["w"] + "}val") is None:
                    shade.set("{" + NS["w"] + "}val", "clear")
            cell_props[:] = sorted(cell_props, key=lambda node: tc_order.get(ET.QName(node).localname, 100))
        picture_paragraph = next(p for p in body.findall("w:p", NS) if p.find(".//wp:inline", NS) is not None)
        props = picture_paragraph.find("w:pPr", NS)
        if props is not None:
            for style in props.findall("w:pStyle", NS):
                props.remove(style)
        inline = picture_paragraph.find(".//wp:inline", NS)
        extent = inline.find("wp:extent", NS)
        width, height = int(extent.get("cx")), int(extent.get("cy"))
        new_width = 7620000  # 600 pt, deliberately wider than the portrait text area.
        new_height = round(height * new_width / width)
        for size in [extent, inline.find(".//pic:spPr/a:xfrm/a:ext", NS)]:
            size.set("cx", str(new_width))
            size.set("cy", str(new_height))

        table_caption = next(p for p in body.findall("w:p", NS) if "".join(p.itertext()).strip().startswith("表4.1"))
        table = table_caption.getnext()
        assert table.tag == "{" + NS["w"] + "}tbl"
        props = table.find("w:tblPr", NS)
        for name in ["tblStyle", "tblW", "tblLayout", "tblInd"]:
            for node in props.findall("w:" + name, NS):
                props.remove(node)
        q = lambda name: "{" + NS["w"] + "}" + name
        grid = table.findall("w:tblGrid/w:gridCol", NS)
        widths = [3000] * len(grid)
        props.insert(0, ET.Element(q("tblW"), {q("type"): "dxa", q("w"): str(sum(widths))}))
        layout = ET.Element(q("tblLayout"), {q("type"): "fixed"})
        following = next((node for node in props if ET.QName(node).localname in
                          {"tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"}), None)
        props.insert(props.index(following) if following is not None else len(props), layout)
        for col, width in zip(grid, widths):
            col.set(q("w"), str(width))
        for row in table.findall("w:tr", NS):
            cells = row.findall("w:tc", NS)
            assert len(cells) == len(widths)
            for cell, width in zip(cells, widths):
                cell_width = cell.find("w:tcPr/w:tcW", NS)
                cell_width.set(q("type"), "dxa")
                cell_width.set(q("w"), str(width))
        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, "w", ZIP_DEFLATED) as destination:
            for part in package.infolist():
                destination.writestr(part, ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                                     if part.filename == "word/document.xml" else package.read(part.filename))
    print(output)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    make_fixture(parser.parse_args().output)
