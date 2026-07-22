"""Parse OneNote XML into markdown and structured data."""

import html
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
    """Parse hierarchy XML into notebook list.

    GetHierarchy returns a different root element depending on the start node:
    an empty start node yields <one:Notebooks>, but scoping to a notebook ID
    yields a bare <one:Notebook>. Handle both so notebook-scoped calls work.
    """
    root = ET.fromstring(xml_str)
    if _local_tag(root.tag) == "Notebook":
        return [_parse_notebook(root)]
    return [_parse_notebook(nb) for nb in root.findall("one:Notebook", NS)]


def _parse_notebook(nb) -> NotebookInfo:
    """Build a NotebookInfo from a <one:Notebook> element."""
    notebook = NotebookInfo(
        id=nb.get("ID", ""),
        name=nb.get("name", ""),
        path=nb.get("path"),
        last_modified=nb.get("lastModifiedTime"),
    )
    notebook.sections = _parse_sections(nb)
    notebook.section_groups = _parse_section_groups(nb)
    return notebook


def parse_section(xml_str: str) -> SectionInfo | None:
    """Parse a section-scoped hierarchy XML (root is a bare <one:Section>).

    Returns None if the XML is not section-scoped.
    """
    root = ET.fromstring(xml_str)
    if _local_tag(root.tag) != "Section":
        return None
    section = SectionInfo(
        id=root.get("ID", ""),
        name=root.get("name", ""),
        path=root.get("path"),
    )
    section.pages = _parse_pages(root)
    return section


def _parse_sections(parent) -> list[SectionInfo]:
    """Parse Section elements under a parent node."""
    sections = []
    for sec in parent.findall("one:Section", NS):
        section = SectionInfo(
            id=sec.get("ID", ""),
            name=sec.get("name", ""),
            path=sec.get("path"),
        )
        section.pages = _parse_pages(sec)
        sections.append(section)
    return sections


def _parse_pages(parent) -> list[PageInfo]:
    """Parse Page elements directly under a parent node."""
    return [
        PageInfo(
            id=page.get("ID", ""),
            name=page.get("name", ""),
            last_modified=page.get("lastModifiedTime"),
            level=int(page.get("pageLevel", "0")),
        )
        for page in parent.findall("one:Page", NS)
    ]


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
    quick_styles = parse_quick_styles(root)

    # Process all Outline elements (main content containers)
    for outline in root.findall(".//one:Outline", NS):
        outline_lines, outline_images, img_counter = _process_outline(
            outline, images_start_index=img_counter, quick_styles=quick_styles
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


def _process_outline(
    outline, images_start_index: int = 0, quick_styles: dict[str, str] | None = None
) -> tuple[list[str], list[ImageRef], int]:
    """Process an Outline element into markdown lines.

    Walks the OE/OEChildren tree explicitly rather than using iter(), so that
    table cells are rendered once (by _process_table) instead of a second time
    as loose text, and so nesting depth survives as markdown indentation.
    """
    lines: list[str] = []
    images: list[ImageRef] = []
    img_counter = images_start_index
    styles = quick_styles or {}

    def walk(oe_children, depth: int) -> None:
        nonlocal img_counter
        for oe in oe_children.findall("one:OE", NS):
            heading = _style_prefix(oe, styles)
            prefix = heading or _list_prefix(oe)
            indent = "" if heading else "    " * depth
            for child in oe:
                tag = _local_tag(child.tag)

                if tag == "T":
                    text = _clean_text(child.text or "")
                    if text.strip():
                        if heading:
                            # OneNote materialises a heading style's weight as an
                            # inline bold span; "# **Title**" would be redundant.
                            text = text.replace("**", "")
                        lines.append(f"{indent}{prefix}{text}")
                        # Only the first line of an OE carries the marker.
                        prefix = heading = ""

                elif tag == "Table":
                    lines.extend(_process_table(child))

                elif tag == "Image":
                    if _get_callback_id(child):
                        img_counter += 1
                        ref = _make_image_ref(child, img_counter)
                        if ref:
                            images.append(ref)
                            lines.append(f"{indent}[Image {ref.index}]")

                elif tag == "InsertedFile":
                    name = child.get("preferredName", "file")
                    lines.append(f"{indent}[Attached: {name}]")

                elif tag == "OEChildren":
                    walk(child, depth + 1)

    for oe_children in outline.findall("one:OEChildren", NS):
        walk(oe_children, 0)

    return lines, images, img_counter


def _style_prefix(oe, quick_styles: dict[str, str]) -> str:
    """Map a OneNote quick style (h1..h6) onto a markdown heading prefix."""
    name = quick_styles.get(oe.get("quickStyleIndex", ""), "")
    match = re.fullmatch(r"h([1-6])", name)
    return "#" * int(match.group(1)) + " " if match else ""


def _list_prefix(oe) -> str:
    """Render a OneNote bullet or number list marker as markdown."""
    lst = oe.find("one:List", NS)
    if lst is None:
        return ""
    if lst.find("one:Bullet", NS) is not None:
        return "- "
    if lst.find("one:Number", NS) is not None:
        return "1. "
    return ""


def parse_quick_styles(root) -> dict[str, str]:
    """Map quickStyleIndex -> style name (e.g. "1" -> "h1") for a page."""
    return {
        qs.get("index", ""): qs.get("name", "")
        for qs in root.findall("one:QuickStyleDef", NS)
    }


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


# OneNote normalises the markup it stores: attributes come back single-quoted
# (style='font-weight:bold'), some are unquoted (lang=ja), and tags may contain
# newlines. These patterns accept either quoting style.
_LINK_RE = re.compile(r"""<a\b[^>]*\bhref=(["'])(.*?)\1[^>]*>(.*?)</a>""", re.I | re.S)
_BOLD_SPAN_RE = re.compile(
    r"""<span\b[^>]*\bstyle=(["'])[^"']*font-weight:\s*bold[^"']*\1[^>]*>(.*?)</span>""",
    re.I | re.S,
)
_ITALIC_SPAN_RE = re.compile(
    r"""<span\b[^>]*\bstyle=(["'])[^"']*font-style:\s*italic[^"']*\1[^>]*>(.*?)</span>""",
    re.I | re.S,
)


def _clean_text(text: str) -> str:
    """Convert OneNote's inline HTML (inside CDATA) to markdown.

    Links, bold and italic are preserved rather than discarded; everything else
    is stripped. Entities are decoded last, in one pass, so that escaped markup
    such as "&amp;lt;" does not get decoded twice into a real tag.
    """
    text = _LINK_RE.sub(lambda m: f"[{m.group(3)}]({m.group(2)})", text)
    text = _BOLD_SPAN_RE.sub(r"**\2**", text)
    text = _ITALIC_SPAN_RE.sub(r"*\2*", text)
    text = re.sub(r"</?(?:b|strong)\b[^>]*>", "**", text, flags=re.I)
    text = re.sub(r"</?(?:i|em)\b[^>]*>", "*", text, flags=re.I)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    # Drop any remaining markup, then decode entities in a single pass.
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(text)


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
