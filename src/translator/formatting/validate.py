from __future__ import annotations

from lxml import etree


def assert_text_unchanged(before: etree._Element, after: etree._Element) -> None:
    before_text = before.xpath(
        ".//w:t/text()",
        namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
    )
    after_text = after.xpath(
        ".//w:t/text()",
        namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
    )
    if before_text != after_text:
        raise ValueError("Formatter changed text content")
