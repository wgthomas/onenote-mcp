"""OneNote COM automation via PowerShell bridge.

64-bit Python can't directly call 32-bit OneNote COM methods (typelib
not registered for Win64). PowerShell handles this through .NET interop.
We shell out to powershell.exe for each COM operation.
"""

import json
import subprocess
import tempfile
import os

# OneNote hierarchy scope constants (HierarchyScope enum)
SELF = 0        # hsSelf — just the node itself
CHILDREN = 1    # hsChildren — node + immediate children
NOTEBOOKS = 2   # hsNotebooks — all notebooks
SECTIONS = 3    # hsSections — notebooks + sections
PAGES = 4       # hsPages — notebooks + sections + pages


def _run_ps(script: str, timeout: int = 30) -> str:
    """Execute a PowerShell script and return stdout.

    For large outputs (XML), we write to a temp file to avoid
    stdout encoding/size issues.
    """
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True,
        text=True,
        timeout=timeout,
        encoding="utf-8",
    )
    if result.returncode != 0:
        stderr = result.stderr.strip()
        raise RuntimeError(f"PowerShell error: {stderr}")
    return result.stdout


def _run_ps_to_file(script: str, timeout: int = 30) -> str:
    """Execute PowerShell script that writes output to a temp file, return contents."""
    tmp = tempfile.mktemp(suffix=".xml")
    try:
        # The script should write to $outFile
        full_script = f'$outFile = "{tmp}"\n{script}'
        _run_ps(full_script, timeout)
        with open(tmp, "r", encoding="utf-8-sig") as f:
            return f.read()
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def get_hierarchy(start_node_id: str = "", scope: int = PAGES) -> str:
    """Get OneNote hierarchy XML."""
    # Escape single quotes in the ID
    safe_id = start_node_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$xml = ""
$onenote.GetHierarchy('{safe_id}', {scope}, [ref]$xml, 2)
$xml | Out-File -FilePath $outFile -Encoding UTF8 -NoNewline
"""
    return _run_ps_to_file(script)


def get_page_content(page_id: str) -> str:
    """Get page content as OneNote XML."""
    safe_id = page_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$xml = ""
$onenote.GetPageContent('{safe_id}', [ref]$xml, 0, 2)
$xml | Out-File -FilePath $outFile -Encoding UTF8 -NoNewline
"""
    return _run_ps_to_file(script)


def get_binary_content(page_id: str, callback_id: str) -> str:
    """Get binary content (image) as base64 string."""
    safe_pid = page_id.replace("'", "''")
    safe_cid = callback_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$b64 = ""
$onenote.GetBinaryPageContent('{safe_pid}', '{safe_cid}', [ref]$b64)
$b64 | Out-File -FilePath $outFile -Encoding UTF8 -NoNewline
"""
    return _run_ps_to_file(script, timeout=60)


def find_pages(query: str, start_node_id: str = "") -> str:
    """Full-text search across notebooks using Windows Search."""
    safe_id = start_node_id.replace("'", "''")
    safe_q = query.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$xml = ""
$onenote.FindPages('{safe_id}', '{safe_q}', [ref]$xml, $false, $false, 2)
$xml | Out-File -FilePath $outFile -Encoding UTF8 -NoNewline
"""
    return _run_ps_to_file(script)


def update_page_content(xml_content: str) -> None:
    """Update/create page content."""
    # Write XML to temp file to avoid quoting issues
    tmp = tempfile.mktemp(suffix=".xml")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(xml_content)
        script = f"""
$onenote = New-Object -ComObject OneNote.Application
$xml = Get-Content -Path '{tmp}' -Raw -Encoding UTF8
$onenote.UpdatePageContent($xml)
"""
        _run_ps(script)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def create_new_page(section_id: str) -> str:
    """Create a new blank page in a section, return the new page ID."""
    safe_id = section_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$pageId = ""
$onenote.CreateNewPage('{safe_id}', [ref]$pageId, 0)
$pageId
"""
    return _run_ps(script).strip()


def navigate_to(object_id: str) -> None:
    """Open an object in the OneNote UI."""
    safe_id = object_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$onenote.NavigateTo('{safe_id}')
"""
    _run_ps(script)


def open_hierarchy(path: str, relative_to_id: str = "") -> str:
    """Create or open a notebook/section via file path. Returns the object ID.

    For sections: pass path to a .one file under a notebook folder.
    For section groups: pass path to a folder under a notebook folder.
    The parent notebook must already exist.
    """
    safe_path = path.replace("'", "''")
    safe_rel = relative_to_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$objectId = ""
$onenote.OpenHierarchy('{safe_path}', '{safe_rel}', [ref]$objectId, 0)
$objectId
"""
    return _run_ps(script).strip()


def delete_hierarchy(object_id: str) -> None:
    """Delete a page, section, or section group."""
    safe_id = object_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$onenote.DeleteHierarchy('{safe_id}')
"""
    _run_ps(script)
