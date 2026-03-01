"""Check current notebook state."""
import sys
sys.path.insert(0, ".")
from onenote_lib import com_client
from onenote_lib.xml_parser import parse_notebooks

# Override timeout for hierarchy call
import onenote_lib.com_client as cc
original = cc._run_ps_to_file

def patched(script, timeout=120):
    return original(script, timeout)

cc._run_ps_to_file = patched

xml = com_client.get_hierarchy("", com_client.PAGES)
notebooks = parse_notebooks(xml)

for nb in notebooks:
    print(f"=== {nb.name} ===")
    total = 0
    for sec in nb.sections:
        pc = len(sec.pages)
        total += pc
        print(f"  [{pc:2d} pages] {sec.name}")
    print(f"  TOTAL: {total} pages")
    print()
