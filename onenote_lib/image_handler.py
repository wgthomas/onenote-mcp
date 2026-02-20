"""Extract and process images from OneNote pages via COM."""

import base64
import io

from PIL import Image

from . import com_client
from .config import config
from .xml_parser import ImageRef


def get_image_base64(page_id: str, callback_id: str, max_size_kb: int | None = None) -> tuple[str, str]:
    """Get a single image as base64, optionally resized.

    Args:
        page_id: OneNote page ID
        callback_id: Image callback ID from page XML
        max_size_kb: Max size in KB (None = use config default)

    Returns:
        Tuple of (base64_data, media_type)
    """
    max_kb = max_size_kb or config.max_image_size_kb
    raw_b64 = com_client.get_binary_content(page_id, callback_id)
    img_bytes = base64.b64decode(raw_b64)

    # Detect format
    img = Image.open(io.BytesIO(img_bytes))
    media_type = _media_type(img.format)

    # Resize if over limit
    if len(img_bytes) > max_kb * 1024:
        img_bytes = _resize_to_limit(img, max_kb)
        # Re-detect after resize (we always output PNG for resized)
        media_type = "image/png"

    return base64.b64encode(img_bytes).decode("ascii"), media_type


def get_all_images(
    page_id: str,
    image_refs: list[ImageRef],
    max_images: int | None = None,
    max_size_kb: int | None = None,
) -> list[dict]:
    """Get all images from a page.

    Args:
        page_id: OneNote page ID
        image_refs: List of ImageRef from xml_parser
        max_images: Max number of images to extract
        max_size_kb: Max size per image in KB

    Returns:
        List of dicts with keys: index, callback_id, base64, media_type
    """
    limit = max_images or config.max_images_per_page
    results = []

    for ref in image_refs[:limit]:
        try:
            b64, mtype = get_image_base64(page_id, ref.callback_id, max_size_kb)
            results.append({
                "index": ref.index,
                "callback_id": ref.callback_id,
                "base64": b64,
                "media_type": mtype,
            })
        except Exception as e:
            results.append({
                "index": ref.index,
                "callback_id": ref.callback_id,
                "error": str(e),
            })

    return results


def _resize_to_limit(img: Image.Image, max_kb: int) -> bytes:
    """Resize image to fit within max_kb."""
    target_bytes = max_kb * 1024

    # Try progressively smaller sizes
    for scale in [0.75, 0.5, 0.35, 0.25, 0.15]:
        new_w = int(img.width * scale)
        new_h = int(img.height * scale)
        if new_w < 100 or new_h < 100:
            break
        resized = img.resize((new_w, new_h), Image.LANCZOS)
        buf = io.BytesIO()
        resized.save(buf, format="PNG", optimize=True)
        if buf.tell() <= target_bytes:
            return buf.getvalue()

    # Last resort: small thumbnail
    img.thumbnail((400, 400), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def _media_type(pil_format: str | None) -> str:
    """Convert PIL format string to MIME type."""
    mapping = {
        "PNG": "image/png",
        "JPEG": "image/jpeg",
        "GIF": "image/gif",
        "BMP": "image/bmp",
        "TIFF": "image/tiff",
        "WEBP": "image/webp",
    }
    return mapping.get(pil_format or "", "image/png")
