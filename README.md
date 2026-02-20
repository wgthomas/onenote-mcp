# OneNote MCP Server

A Model Context Protocol (MCP) server that gives Claude (or any MCP client) direct access to local Microsoft OneNote notebooks via the COM API. No Azure registration, no API keys, no cloud auth — just your local OneNote desktop app.

**Windows only.** Requires OneNote desktop installed and running.

## Features

- **13 tools** for full OneNote access
- Navigate notebooks, sections, section groups, and pages
- Read pages as clean markdown with image references
- Full-text search across all notebooks (via Windows Search)
- Extract embedded images as native MCP Image objects (Claude can see them directly)
- Optional vision analysis of images via any OpenAI-compatible vision API (LM Studio, ollama, etc.)
- Create new pages with HTML content

## Tools

| Category | Tool | Description |
|----------|------|-------------|
| Navigation | `onenote_list_notebooks` | List all open notebooks |
| | `onenote_list_sections` | List sections in a notebook |
| | `onenote_list_pages` | List pages in a section |
| | `onenote_get_notebook_tree` | Full hierarchy tree |
| Content | `onenote_get_page` | Page as clean markdown |
| | `onenote_get_page_raw` | Raw OneNote XML |
| | `onenote_get_page_images` | All images as MCP Image objects |
| | `onenote_get_image` | Single image by callback ID |
| Search | `onenote_search` | Full-text search (all notebooks) |
| | `onenote_search_in_notebook` | Search within one notebook |
| Vision | `onenote_analyze_page_visuals` | Vision model analysis of all page images |
| | `onenote_describe_image` | Vision model analysis of one image |
| Write | `onenote_create_page` | Create a new page |

## Installation

```bash
pip install -r requirements.txt
```

Dependencies: `mcp[cli]`, `pydantic-settings`, `Pillow`, `httpx`

## Configuration

Add to your Claude Code MCP config (`~/.claude/mcp.json`):

```json
{
  "mcpServers": {
    "onenote": {
      "type": "stdio",
      "command": "python",
      "args": ["path/to/onenote-mcp/onenote_mcp.py"],
      "env": {}
    }
  }
}
```

### Vision (Optional)

To enable image analysis via a local vision model (e.g., LM Studio):

```json
"env": {
  "ONENOTE_VISION_URL": "http://localhost:1234",
  "ONENOTE_VISION_MODEL": "your-vision-model-name",
  "ONENOTE_VISION_FALLBACK_URL": "http://other-host:1234",
  "ONENOTE_VISION_FALLBACK_MODEL": "fallback-model-name"
}
```

Any OpenAI-compatible `/v1/chat/completions` endpoint with vision support works.

## How It Works

64-bit Python can't call 32-bit OneNote COM objects directly. This server bridges the gap by shelling out to PowerShell for every COM operation, which handles the COM interop natively. Results are passed back via temp files to avoid stdout encoding issues.

## Requirements

- Windows 10/11
- Microsoft OneNote desktop (not the UWP/Store version)
- Python 3.11+
- OneNote must be running (or at least installed — COM will launch it)

## License

MIT
