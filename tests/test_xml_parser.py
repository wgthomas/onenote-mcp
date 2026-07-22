"""Unit tests for OneNote XML parser."""

import pytest

from onenote_lib.xml_parser import (
    parse_notebooks,
    parse_page_to_markdown,
    parse_search_results,
    parse_section,
)

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


HIERARCHY_XML = f"""<?xml version="1.0"?>
<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="Work Notes" ID="nb-001" path="C:\\Users\\test\\Work Notes"
                lastModifiedTime="2026-02-15T10:00:00Z">
    <one:Section name="Meeting Notes" ID="sec-001" path="C:\\Users\\test\\Work Notes\\Meeting Notes.one">
      <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z" pageLevel="0"/>
      <one:Page ID="page-002" name="Sprint Review" lastModifiedTime="2026-02-13T14:00:00Z" pageLevel="0"/>
    </one:Section>
    <one:SectionGroup name="Archive" ID="sg-001">
      <one:Section name="Old Notes" ID="sec-002">
        <one:Page ID="page-003" name="Archived Page" pageLevel="0"/>
      </one:Section>
    </one:SectionGroup>
    <one:SectionGroup name="Recycle Bin" ID="sg-bin" isRecycleBin="true">
      <one:Section name="Deleted" ID="sec-deleted"/>
    </one:SectionGroup>
  </one:Notebook>
  <one:Notebook name="Personal" ID="nb-002" path="C:\\Users\\test\\Personal"
                lastModifiedTime="2026-02-10T08:00:00Z">
    <one:Section name="Journal" ID="sec-003">
      <one:Page ID="page-004" name="Feb 10" pageLevel="0"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>"""


PAGE_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-001" name="Test Page">
  <one:Title>
    <one:OE><one:T><![CDATA[Test Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[This is paragraph one.]]></one:T>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[This is paragraph two with <b>bold</b> text.]]></one:T>
      </one:OE>
      <one:OE>
        <one:Image>
          <one:Size width="640" height="480" isSetByUser="true"/>
          <one:CallbackID callbackID="img-001"/>
        </one:Image>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[Text after image.]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:Table>
          <one:Row>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 1]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 2]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>
          <one:Row>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 1]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 2]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>
        </one:Table>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


PAGE_NO_IMAGES_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-005" name="Plain Page">
  <one:Title>
    <one:OE><one:T><![CDATA[Plain Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[Just text, no images.]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


SEARCH_XML = f"""<?xml version="1.0"?>
<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="Work Notes" ID="nb-001">
    <one:Section name="Meeting Notes" ID="sec-001">
      <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>"""


class TestParseNotebooks:
    def test_basic_parsing(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        assert len(notebooks) == 2
        assert notebooks[0].name == "Work Notes"
        assert notebooks[0].id == "nb-001"
        assert notebooks[1].name == "Personal"

    def test_sections(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        nb = notebooks[0]
        assert len(nb.sections) == 1
        assert nb.sections[0].name == "Meeting Notes"
        assert len(nb.sections[0].pages) == 2

    def test_section_groups(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        nb = notebooks[0]
        # Should have Archive but NOT Recycle Bin
        assert len(nb.section_groups) == 1
        assert nb.section_groups[0].name == "Archive"
        assert len(nb.section_groups[0].sections) == 1

    def test_pages(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        pages = notebooks[0].sections[0].pages
        assert len(pages) == 2
        assert pages[0].name == "Monday Standup"
        assert pages[0].id == "page-001"

    def test_nested_section_group_pages(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        sg = notebooks[0].section_groups[0]
        assert sg.sections[0].pages[0].name == "Archived Page"


class TestParsePageToMarkdown:
    def test_basic_content(self):
        md, images = parse_page_to_markdown(PAGE_XML)
        assert "# Test Page" in md
        assert "This is paragraph one." in md
        # Bold is carried through as markdown rather than stripped.
        assert "This is paragraph two with **bold** text." in md
        assert "Text after image." in md

    def test_image_references(self):
        md, images = parse_page_to_markdown(PAGE_XML)
        assert len(images) == 1
        assert images[0].callback_id == "img-001"
        assert images[0].width == 640.0
        assert images[0].height == 480.0
        assert "[Image 1]" in md

    def test_table_parsing(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        assert "Header 1" in md
        assert "Header 2" in md
        assert "Data 1" in md
        assert "|" in md
        assert "---" in md

    def test_no_images(self):
        md, images = parse_page_to_markdown(PAGE_NO_IMAGES_XML)
        assert len(images) == 0
        assert "Just text, no images." in md

    def test_html_stripping(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        # <b> tags should be stripped
        assert "<b>" not in md
        assert "bold" in md


class TestParseSearchResults:
    def test_basic_search(self):
        results = parse_search_results(SEARCH_XML)
        assert len(results) == 1
        assert results[0]["page_name"] == "Monday Standup"
        assert results[0]["notebook"] == "Work Notes"
        assert results[0]["section"] == "Meeting Notes"

    def test_empty_search(self):
        empty_xml = f'<one:Notebooks xmlns:one="{NS}"/>'
        results = parse_search_results(empty_xml)
        assert len(results) == 0


# GetHierarchy returns a different root element depending on the start node:
# "" yields <one:Notebooks>, a notebook ID yields a bare <one:Notebook>, and a
# section ID yields a bare <one:Section> whose pages are direct children.

SCOPED_NOTEBOOK_XML = f"""<?xml version="1.0"?>
<one:Notebook xmlns:one="{NS}" name="Work Notes" ID="nb-001">
  <one:Section name="Meeting Notes" ID="sec-001">
    <one:Page ID="page-001" name="Monday Standup" pageLevel="0"/>
  </one:Section>
  <one:SectionGroup name="Archive" ID="sg-001">
    <one:Section name="Old Notes" ID="sec-002"/>
  </one:SectionGroup>
</one:Notebook>"""

SCOPED_SECTION_XML = f"""<?xml version="1.0"?>
<one:Section xmlns:one="{NS}" name="Meeting Notes" ID="sec-001">
  <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z" pageLevel="0"/>
  <one:Page ID="page-002" name="Sprint Review" pageLevel="1"/>
</one:Section>"""


class TestScopedHierarchy:
    def test_bare_notebook_root(self):
        notebooks = parse_notebooks(SCOPED_NOTEBOOK_XML)
        assert len(notebooks) == 1
        assert notebooks[0].name == "Work Notes"
        assert notebooks[0].sections[0].name == "Meeting Notes"
        assert notebooks[0].section_groups[0].name == "Archive"

    def test_bare_section_root(self):
        section = parse_section(SCOPED_SECTION_XML)
        assert section is not None
        assert section.name == "Meeting Notes"
        assert [p.name for p in section.pages] == ["Monday Standup", "Sprint Review"]
        assert section.pages[1].level == 1

    def test_parse_section_rejects_other_roots(self):
        assert parse_section(SCOPED_NOTEBOOK_XML) is None
        assert parse_section(HIERARCHY_XML) is None


TABLE_PAGE_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-t" name="Table Page">
  <one:Outline><one:OEChildren>
    <one:OE><one:Table>
      <one:Row>
        <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[H1]]></one:T></one:OE></one:OEChildren></one:Cell>
        <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[H2]]></one:T></one:OE></one:OEChildren></one:Cell>
      </one:Row>
    </one:Table></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""

RICH_PAGE_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-r" name="Rich Page">
  <one:QuickStyleDef index="0" name="PageTitle"/>
  <one:QuickStyleDef index="1" name="p"/>
  <one:QuickStyleDef index="2" name="h2"/>
  <one:Outline><one:OEChildren>
    <one:OE quickStyleIndex="2"><one:T><![CDATA[A Heading]]></one:T></one:OE>
    <one:OE quickStyleIndex="1"><one:T><![CDATA[top level]]></one:T>
      <one:OEChildren>
        <one:OE><one:T><![CDATA[nested once]]></one:T>
          <one:OEChildren><one:OE><one:T><![CDATA[nested twice]]></one:T></one:OE></one:OEChildren>
        </one:OE>
      </one:OEChildren>
    </one:OE>
    <one:OE><one:T><![CDATA[see <a href="https://example.com/x">the docs</a>]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[<span style="font-weight:bold">strong</span>]]></one:T></one:OE>
    <one:OE><one:T><![CDATA[5 &amp;lt; 6]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""


class TestMarkdownFidelity:
    def test_table_cells_are_not_duplicated(self):
        md, _ = parse_page_to_markdown(TABLE_PAGE_XML)
        # Each cell renders once, inside the markdown table — not again as loose text.
        assert md.count("H1") == 1
        assert "| H1 | H2 |" in md

    def test_quick_style_becomes_heading(self):
        md, _ = parse_page_to_markdown(RICH_PAGE_XML)
        assert "## A Heading" in md

    def test_nesting_becomes_indentation(self):
        md, _ = parse_page_to_markdown(RICH_PAGE_XML)
        assert "\ntop level" in md
        assert "\n    nested once" in md
        assert "\n        nested twice" in md

    def test_links_keep_their_url(self):
        md, _ = parse_page_to_markdown(RICH_PAGE_XML)
        assert "[the docs](https://example.com/x)" in md

    def test_bold_span_becomes_markdown(self):
        md, _ = parse_page_to_markdown(RICH_PAGE_XML)
        assert "**strong**" in md

    def test_onenote_normalised_markup(self):
        # What OneNote actually stores after it rewrites a page: single-quoted
        # attributes, unquoted lang attributes, and newlines inside tags.
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p" name="Normalised">
  <one:Outline><one:OEChildren><one:OE><one:T><![CDATA[<span
lang=ja>item </span><span style='font-weight:bold' lang=x-none>bold</span><span
lang=ja> and </span><a href="https://example.com"><span lang=x-none>link</span></a>]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "**bold**" in md
        assert "[link](https://example.com)" in md

    def test_heading_does_not_double_up_bold(self):
        # OneNote renders a heading style's weight as an inline bold span.
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p" name="H">
  <one:QuickStyleDef index="0" name="h1"/>
  <one:Outline><one:OEChildren>
    <one:OE quickStyleIndex="0"><one:T><![CDATA[<span style='font-weight:bold'>Title</span>]]></one:T></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "# Title" in md
        assert "**" not in md

    def test_entities_are_decoded_once(self):
        md, _ = parse_page_to_markdown(RICH_PAGE_XML)
        # "&amp;lt;" is an escaped "&lt;" — it must not decode all the way to "<".
        assert "5 &lt; 6" in md
