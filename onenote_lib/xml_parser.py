"""Parse OneNote XML into markdown and structured data."""

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

# OneNote 2013 XML namespace
NS = {"one": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


@dataclass
class ImageRef:
    """Reference to an image in a OneNote page."""
    callback_id: str
    index: int
    width: float | None = None
    height: float | None = None
    alt_text: str | None = None


@dataclass
class PageInfo:
    """Parsed page metadata."""
    id: str
    name: str
    last_modified: str | None = None
    level: int = 0


@dataclass
class SectionInfo:
    """Parsed section metadata."""
    id: str
    name: str
    path: str | None = None
    pages: list[PageInfo] = field(default_factory=list)


@dataclass
class SectionGroupInfo:
    """Parsed section group metadata."""
    id: str
    name: str
    sections: list[SectionInfo] = field(default_factory=list)
    section_groups: list["SectionGroupInfo"] = field(default_factory=list)


@dataclass
class NotebookInfo:
    """Parsed notebook metadata."""
    id: str
    name: str
    path: str | None = None
    last_modified: str | None = None
    sections: list[SectionInfo] = field(default_factory=list)
    section_groups: list[SectionGroupInfo] = field(default_factory=list)


def parse_notebooks(xml_str: str) -> list[NotebookInfo]:
    """Parse hierarchy XML into notebook list."""
    root = ET.fromstring(xml_str)
    notebooks = []
    for nb in root.findall("one:Notebook", NS):
        notebook = NotebookInfo(
            id=nb.get("ID", ""),
            name=nb.get("name", ""),
            path=nb.get("path"),
            last_modified=nb.get("lastModifiedTime"),
        )
        notebook.sections = _parse_sections(nb)
        notebook.section_groups = _parse_section_groups(nb)
        notebooks.append(notebook)
    return notebooks


def _parse_sections(parent) -> list[SectionInfo]:
    """Parse Section elements under a parent node."""
    sections = []
    for sec in parent.findall("one:Section", NS):
        section = SectionInfo(
            id=sec.get("ID", ""),
            name=sec.get("name", ""),
            path=sec.get("path"),
        )
        for page in sec.findall("one:Page", NS):
            section.pages.append(PageInfo(
                id=page.get("ID", ""),
                name=page.get("name", ""),
                last_modified=page.get("lastModifiedTime"),
                level=int(page.get("pageLevel", "0")),
            ))
        sections.append(section)
    return sections


def _parse_section_groups(parent) -> list[SectionGroupInfo]:
    """Parse SectionGroup elements recursively."""
    groups = []
    for sg in parent.findall("one:SectionGroup", NS):
        # Skip recycle bin
        if sg.get("isRecycleBin") == "true":
            continue
        group = SectionGroupInfo(
            id=sg.get("ID", ""),
            name=sg.get("name", ""),
        )
        group.sections = _parse_sections(sg)
        group.section_groups = _parse_section_groups(sg)
        groups.append(group)
    return groups


def parse_page_to_markdown(xml_str: str) -> tuple[str, list[ImageRef]]:
    """Convert OneNote page XML to markdown text + image references.

    Returns:
        Tuple of (markdown_text, list_of_image_refs)
    """
    root = ET.fromstring(xml_str)
    title = root.get("name", root.get("ID", "Untitled"))
    lines = [f"# {title}", ""]

    images: list[ImageRef] = []
    img_counter = 0

    # Process all Outline elements (main content containers)
    for outline in root.findall(".//one:Outline", NS):
        outline_lines, outline_images, img_counter = _process_outline(
            outline, images_start_index=img_counter
        )
        lines.extend(outline_lines)
        images.extend(outline_images)
        lines.append("")

    # Process top-level images (outside outlines)
    for img in root.findall(".//one:Image", NS):
        # Skip images already found inside outlines
        cb_id = _get_callback_id(img)
        if cb_id and not any(i.callback_id == cb_id for i in images):
            img_counter += 1
            ref = _make_image_ref(img, img_counter)
            if ref:
                images.append(ref)
                lines.append(f"[Image {ref.index}]")

    return "\n".join(lines).strip(), images


def _process_outline(outline, images_start_index: int = 0) -> tuple[list[str], list[ImageRef], int]:
    """Process an Outline element into markdown lines."""
    lines = []
    images = []
    img_counter = images_start_index

    for oe in outline.iter():
        tag = _local_tag(oe.tag)

        if tag == "T":
            # Text element — extract CDATA content
            text = oe.text or ""
            text = _clean_text(text)
            if text.strip():
                lines.append(text)

        elif tag == "Image":
            cb_id = _get_callback_id(oe)
            if cb_id:
                img_counter += 1
                ref = _make_image_ref(oe, img_counter)
                if ref:
                    images.append(ref)
                    lines.append(f"[Image {ref.index}]")

        elif tag == "Table":
            table_lines = _process_table(oe)
            lines.extend(table_lines)

        elif tag == "InsertedFile":
            name = oe.get("preferredName", "file")
            lines.append(f"[Attached: {name}]")

    return lines, images, img_counter


def _process_table(table_elem) -> list[str]:
    """Convert a OneNote table to markdown table."""
    rows = table_elem.findall("one:Row", NS)
    if not rows:
        return []

    md_rows = []
    for row in rows:
        cells = row.findall("one:Cell", NS)
        cell_texts = []
        for cell in cells:
            # Collect all text in the cell
            texts = []
            for t in cell.iter():
                if _local_tag(t.tag) == "T" and t.text:
                    texts.append(_clean_text(t.text).strip())
            cell_texts.append(" ".join(texts) if texts else "")
        md_rows.append("| " + " | ".join(cell_texts) + " |")

    if len(md_rows) >= 1:
        # Insert header separator after first row
        col_count = md_rows[0].count("|") - 1
        separator = "| " + " | ".join(["---"] * col_count) + " |"
        md_rows.insert(1, separator)

    return md_rows


def _get_callback_id(img_elem) -> str | None:
    """Extract callbackID from an Image element.

    OneNote stores it as a child element: <one:CallbackID callbackID="..."/>
    not as an attribute on the Image tag itself.
    """
    # Check child element first (actual OneNote format)
    cb_elem = img_elem.find("one:CallbackID", NS)
    if cb_elem is not None:
        return cb_elem.get("callbackID")
    # Fallback: check as attribute (for compatibility)
    return img_elem.get("callbackID")


def _make_image_ref(img_elem, index: int) -> ImageRef | None:
    """Create an ImageRef from an Image element."""
    cb_id = _get_callback_id(img_elem)
    if not cb_id:
        return None

    # Try to get dimensions from Size child
    width = None
    height = None
    size = img_elem.find("one:Size", NS)
    if size is not None:
        w = size.get("width")
        h = size.get("height")
        if w:
            width = float(w)
        if h:
            height = float(h)

    return ImageRef(
        callback_id=cb_id,
        index=index,
        width=width,
        height=height,
    )


def _local_tag(tag: str) -> str:
    """Strip namespace from tag name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _clean_text(text: str) -> str:
    """Clean OneNote text content (strip HTML-like tags from CDATA)."""
    # OneNote sometimes wraps text in span tags with styles
    text = re.sub(r"<[^>]+>", "", text)
    # Decode common HTML entities
    text = text.replace("&amp;", "&")
    text = text.replace("&lt;", "<")
    text = text.replace("&gt;", ">")
    text = text.replace("&quot;", '"')
    text = text.replace("&apos;", "'")
    text = text.replace("&nbsp;", " ")
    return text


def parse_search_results(xml_str: str) -> list[dict]:
    """Parse FindPages result XML into a list of matches."""
    root = ET.fromstring(xml_str)
    results = []

    for nb in root.findall("one:Notebook", NS):
        nb_name = nb.get("name", "")
        for sec in nb.findall(".//one:Section", NS):
            sec_name = sec.get("name", "")
            for page in sec.findall("one:Page", NS):
                results.append({
                    "page_id": page.get("ID", ""),
                    "page_name": page.get("name", ""),
                    "notebook": nb_name,
                    "section": sec_name,
                    "last_modified": page.get("lastModifiedTime", ""),
                })

    return results
