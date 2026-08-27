from __future__ import annotations

from lxml import etree

from translator.formatting.profiles import FormattingProfile

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"


def format_target_paragraph(paragraph: etree._Element, profile: FormattingProfile) -> None:
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        paragraph.insert(0, ppr)
    spacing = ppr.find(W + "spacing")
    if spacing is None:
        spacing = etree.SubElement(ppr, W + "spacing")
    spacing.set(W + "after", str(profile.target_space_after_twips))


def format_existing_heading(paragraph: etree._Element, profile: FormattingProfile) -> bool:
    if not profile.heading_keep_with_next:
        return False
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        return False
    style = ppr.find(W + "pStyle")
    if style is None or "heading" not in style.get(W + "val", "").lower():
        return False
    if ppr.find(W + "keepNext") is None:
        etree.SubElement(ppr, W + "keepNext")
    return True


def apply_document_formatting(root: etree._Element, profile: FormattingProfile) -> dict[str, int]:
    headings = sum(format_existing_heading(p, profile) for p in root.xpath(".//w:p", namespaces=NS))
    tables = 0
    if profile.table_cell_margin_twips is not None:
        for table in root.xpath(".//w:tbl", namespaces=NS):
            tables += 1
            for cell in table.xpath(".//w:tc", namespaces=NS):
                tcpr = cell.find(W + "tcPr")
                if tcpr is None:
                    tcpr = etree.Element(W + "tcPr")
                    cell.insert(0, tcpr)
                margins = tcpr.find(W + "tcMar")
                if margins is None:
                    margins = etree.SubElement(tcpr, W + "tcMar")
                for side in ("top", "start", "bottom", "end"):
                    node = margins.find(W + side)
                    if node is None:
                        node = etree.SubElement(margins, W + side)
                    node.set(W + "w", str(profile.table_cell_margin_twips))
                    node.set(W + "type", "dxa")
    return {"headings_formatted": headings, "tables_formatted": tables}
