"""OneNote MCP Server — COM automation for OneNote desktop.

Provides 13 tools for navigating, reading, searching, and analyzing
OneNote content including embedded images and diagrams.
"""

import json

from mcp.server.fastmcp import FastMCP, Image

from onenote_lib import com_client
from onenote_lib.config import config
from onenote_lib.image_handler import get_all_images, get_image_base64
from onenote_lib.vision import describe_image, describe_images
from onenote_lib.xml_parser import (
    NotebookInfo,
    SectionGroupInfo,
    parse_notebooks,
    parse_page_to_markdown,
    parse_search_results,
)

mcp = FastMCP(
    "OneNote MCP",
    description="Access OneNote desktop notebooks via COM automation. "
    "Read, search, and analyze pages including embedded images.",
)


# ── Navigation Tools ─────────────────────────────────────────────────


@mcp.tool()
def onenote_list_notebooks() -> str:
    """List all open notebooks with their IDs, names, paths, and last modified times."""
    xml = com_client.get_hierarchy("", com_client.NOTEBOOKS)
    notebooks = parse_notebooks(xml)
    result = []
    for nb in notebooks:
        result.append({
            "id": nb.id,
            "name": nb.name,
            "path": nb.path,
            "last_modified": nb.last_modified,
        })
    return json.dumps(result, indent=2)


@mcp.tool()
def onenote_list_sections(notebook_id: str) -> str:
    """List all sections in a notebook, including sections inside section groups.

    Args:
        notebook_id: The notebook's OneNote ID (from onenote_list_notebooks)
    """
    xml = com_client.get_hierarchy(notebook_id, com_client.SECTIONS)
    notebooks = parse_notebooks(xml)
    if not notebooks:
        return json.dumps({"error": "Notebook not found"})

    nb = notebooks[0]
    result = _flatten_sections(nb.sections, nb.section_groups)
    return json.dumps(result, indent=2)


@mcp.tool()
def onenote_list_pages(section_id: str) -> str:
    """List all pages in a section with titles and last modified times.

    Args:
        section_id: The section's OneNote ID (from onenote_list_sections)
    """
    xml = com_client.get_hierarchy(section_id, com_client.PAGES)
    notebooks = parse_notebooks(xml)

    pages = []
    for nb in notebooks:
        for sec in nb.sections:
            if sec.id == section_id:
                for p in sec.pages:
                    pages.append({
                        "id": p.id,
                        "name": p.name,
                        "last_modified": p.last_modified,
                        "level": p.level,
                    })
                return json.dumps(pages, indent=2)
        # Check section groups
        found = _find_section_pages(nb.section_groups, section_id)
        if found is not None:
            return json.dumps(found, indent=2)

    return json.dumps({"error": "Section not found"})


@mcp.tool()
def onenote_get_notebook_tree(notebook_id: str = "") -> str:
    """Get the full hierarchy: notebooks -> section groups -> sections -> page titles.

    Args:
        notebook_id: Optional notebook ID to scope the tree. Empty string = all notebooks.
    """
    xml = com_client.get_hierarchy(notebook_id, com_client.PAGE_INFO)
    notebooks = parse_notebooks(xml)
    result = []
    for nb in notebooks:
        result.append(_notebook_to_tree(nb))
    return json.dumps(result, indent=2)


# ── Content Retrieval Tools ──────────────────────────────────────────


@mcp.tool()
def onenote_get_page(page_id: str) -> str:
    """Get a page's content as clean markdown. Images are listed as [Image N] references
    with callback IDs that can be retrieved with onenote_get_page_images or onenote_get_image.

    Args:
        page_id: The page's OneNote ID (from onenote_list_pages or search)
    """
    xml = com_client.get_page_content(page_id)
    markdown, images = parse_page_to_markdown(xml)

    if images:
        markdown += "\n\n---\n**Image References:**\n"
        for img in images:
            dims = ""
            if img.width and img.height:
                dims = f" ({img.width:.0f}x{img.height:.0f})"
            markdown += f"- [Image {img.index}]: callback_id=`{img.callback_id}`{dims}\n"

    return markdown


@mcp.tool()
def onenote_get_page_raw(page_id: str) -> str:
    """Get a page's raw OneNote XML content for debugging.

    Args:
        page_id: The page's OneNote ID
    """
    return com_client.get_page_content(page_id)


@mcp.tool()
def onenote_get_page_images(
    page_id: str,
    max_images: int = 10,
    max_size_kb: int = 512,
) -> list:
    """Extract all images from a page and return them as viewable images.
    Claude can see these images natively for analysis.

    Args:
        page_id: The page's OneNote ID
        max_images: Maximum number of images to extract (default 10)
        max_size_kb: Maximum size per image in KB (default 512, images are resized if larger)
    """
    xml = com_client.get_page_content(page_id)
    _, image_refs = parse_page_to_markdown(xml)

    if not image_refs:
        return ["No images found on this page."]

    images = get_all_images(page_id, image_refs, max_images, max_size_kb)

    result = []
    for img in images:
        if "error" in img:
            result.append(f"[Image {img['index']}] Error: {img['error']}")
        else:
            result.append(f"[Image {img['index']}] (callback_id: {img['callback_id']})")
            result.append(Image(data=img["base64"], media_type=img["media_type"]))

    return result


@mcp.tool()
def onenote_get_image(
    page_id: str,
    callback_id: str,
    max_size_kb: int = 512,
) -> list:
    """Get a single image by its callback ID. Returns the image for Claude to see natively.

    Args:
        page_id: The page's OneNote ID
        callback_id: The image's callback ID (from onenote_get_page output)
        max_size_kb: Maximum size in KB (default 512, image is resized if larger)
    """
    b64, media_type = get_image_base64(page_id, callback_id, max_size_kb)
    return [Image(data=b64, media_type=media_type)]


# ── Search Tools ─────────────────────────────────────────────────────


@mcp.tool()
def onenote_search(query: str) -> str:
    """Full-text search across all open notebooks. Uses Windows Search indexing.

    Args:
        query: Search query string
    """
    xml = com_client.find_pages(query)
    results = parse_search_results(xml)
    if not results:
        return json.dumps({"message": "No results found", "query": query})
    return json.dumps(results, indent=2)


@mcp.tool()
def onenote_search_in_notebook(notebook_id: str, query: str) -> str:
    """Search within a specific notebook.

    Args:
        notebook_id: The notebook's OneNote ID
        query: Search query string
    """
    xml = com_client.find_pages(query, notebook_id)
    results = parse_search_results(xml)
    if not results:
        return json.dumps({"message": "No results found", "query": query, "notebook_id": notebook_id})
    return json.dumps(results, indent=2)


# ── Vision Analysis Tools ────────────────────────────────────────────


@mcp.tool()
async def onenote_analyze_page_visuals(
    page_id: str,
    prompt: str = "",
    max_images: int = 5,
    max_size_kb: int = 512,
) -> list:
    """Fetch all images from a page, send each to a vision model for description,
    and return both the descriptions and the raw images for Claude to see.

    Requires a vision-capable LLM server (set ONENOTE_VISION_URL and ONENOTE_VISION_MODEL).

    Args:
        page_id: The page's OneNote ID
        prompt: Optional custom prompt for the vision model
        max_images: Maximum number of images to process (default 5)
        max_size_kb: Maximum size per image in KB (default 512)
    """
    xml = com_client.get_page_content(page_id)
    _, image_refs = parse_page_to_markdown(xml)

    if not image_refs:
        return ["No images found on this page."]

    images = get_all_images(page_id, image_refs, max_images, max_size_kb)
    analyzed = await describe_images(images, prompt or None)

    result = []
    for img in analyzed:
        if "error" in img:
            result.append(f"[Image {img['index']}] Error: {img['error']}")
            if "description" in img:
                result.append(f"Description: {img['description']}")
        else:
            result.append(f"[Image {img['index']}] Vision analysis: {img['description']}")
            result.append(Image(data=img["base64"], media_type=img["media_type"]))

    return result


@mcp.tool()
async def onenote_describe_image(
    page_id: str,
    callback_id: str,
    prompt: str = "Describe this image in detail. If it's a diagram, explain the structure and relationships shown.",
    max_size_kb: int = 512,
) -> list:
    """Send a single image to the vision model with a custom prompt.
    Returns both the description and the raw image.

    Args:
        page_id: The page's OneNote ID
        callback_id: The image's callback ID
        prompt: Custom prompt for the vision model
        max_size_kb: Maximum size in KB (default 512)
    """
    b64, media_type = get_image_base64(page_id, callback_id, max_size_kb)
    description = await describe_image(b64, media_type, prompt)

    return [
        f"Vision analysis: {description}",
        Image(data=b64, media_type=media_type),
    ]


# ── Write Tool ───────────────────────────────────────────────────────


@mcp.tool()
def onenote_create_page(section_id: str, title: str, content_html: str = "") -> str:
    """Create a new page in a section.

    Args:
        section_id: The section's OneNote ID where the page will be created
        title: Page title
        content_html: Optional HTML content for the page body.
            Use simple HTML: <p>, <b>, <i>, <ul>/<li>, <table>, <h1>-<h6>.
            Leave empty for a blank page.
    """
    ns = "http://schemas.microsoft.com/office/onenote/2013/onenote"
    try:
        new_page_id = com_client.create_new_page(section_id)

        if content_html or title:
            import xml.etree.ElementTree as ET
            page_xml = com_client.get_page_content(new_page_id)
            root = ET.fromstring(page_xml)
            ns_map = {"one": ns}

            title_elem = root.find(".//one:Title//one:T", ns_map)
            if title_elem is not None:
                title_elem.text = title

            if content_html:
                outline = ET.SubElement(root, f"{{{ns}}}Outline")
                oe_children = ET.SubElement(outline, f"{{{ns}}}OEChildren")
                oe = ET.SubElement(oe_children, f"{{{ns}}}OE")
                t = ET.SubElement(oe, f"{{{ns}}}T")
                t.text = content_html

            updated_xml = ET.tostring(root, encoding="unicode", xml_declaration=True)
            com_client.update_page_content(updated_xml)

        return json.dumps({
            "status": "created",
            "page_id": new_page_id,
            "title": title,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


# ── Helpers ──────────────────────────────────────────────────────────


def _flatten_sections(
    sections: list, section_groups: list, prefix: str = ""
) -> list[dict]:
    """Flatten sections and section groups into a flat list with group paths."""
    result = []
    for sec in sections:
        result.append({
            "id": sec.id,
            "name": sec.name,
            "group": prefix or None,
            "path": sec.path,
            "page_count": len(sec.pages),
        })
    for sg in section_groups:
        group_path = f"{prefix}/{sg.name}" if prefix else sg.name
        result.extend(_flatten_sections(sg.sections, sg.section_groups, group_path))
    return result


def _find_section_pages(section_groups: list[SectionGroupInfo], section_id: str):
    """Recursively find pages in a section within section groups."""
    for sg in section_groups:
        for sec in sg.sections:
            if sec.id == section_id:
                return [
                    {
                        "id": p.id,
                        "name": p.name,
                        "last_modified": p.last_modified,
                        "level": p.level,
                    }
                    for p in sec.pages
                ]
        found = _find_section_pages(sg.section_groups, section_id)
        if found is not None:
            return found
    return None


def _notebook_to_tree(nb: NotebookInfo) -> dict:
    """Convert a NotebookInfo to a tree dict."""
    tree = {
        "id": nb.id,
        "name": nb.name,
        "sections": [],
        "section_groups": [],
    }
    for sec in nb.sections:
        tree["sections"].append({
            "id": sec.id,
            "name": sec.name,
            "pages": [{"id": p.id, "name": p.name, "level": p.level} for p in sec.pages],
        })
    for sg in nb.section_groups:
        tree["section_groups"].append(_section_group_to_tree(sg))
    return tree


def _section_group_to_tree(sg: SectionGroupInfo) -> dict:
    """Convert a SectionGroupInfo to a tree dict."""
    tree = {
        "id": sg.id,
        "name": sg.name,
        "sections": [],
        "section_groups": [],
    }
    for sec in sg.sections:
        tree["sections"].append({
            "id": sec.id,
            "name": sec.name,
            "pages": [{"id": p.id, "name": p.name, "level": p.level} for p in sec.pages],
        })
    for child in sg.section_groups:
        tree["section_groups"].append(_section_group_to_tree(child))
    return tree


if __name__ == "__main__":
    mcp.run()
