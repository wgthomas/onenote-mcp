"""Optional vision analysis via OpenAI-compatible vision API (primary + fallback)."""

import base64

import httpx

from .config import config


async def _call_vision(url: str, model: str, image_b64: str, media_type: str, prompt: str) -> str:
    """Make a vision API call to a specific endpoint."""
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{media_type};base64,{image_b64}"
                        },
                    },
                    {
                        "type": "text",
                        "text": prompt,
                    },
                ],
            }
        ],
        "max_tokens": 1024,
    }

    async with httpx.AsyncClient(timeout=60.0) as client:
        resp = await client.post(f"{url}/v1/chat/completions", json=payload)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]


async def describe_image(
    image_b64: str,
    media_type: str = "image/png",
    prompt: str = "Describe this image in detail. If it's a diagram, explain the structure and relationships shown.",
) -> str:
    """Send an image to the vision model for description.

    Tries Soundwave first, falls back to Megatron2 local LM Studio.

    Args:
        image_b64: Base64-encoded image data
        media_type: MIME type of the image
        prompt: Custom prompt for the vision model

    Returns:
        Description string, or error message if both unavailable
    """
    if not config.vision_model:
        return "[Vision not configured — set ONENOTE_VISION_MODEL env var.]"

    # Primary
    try:
        return await _call_vision(config.vision_url, config.vision_model, image_b64, media_type, prompt)
    except (httpx.ConnectError, httpx.ConnectTimeout):
        pass
    except Exception as e:
        pass  # Fall through to fallback

    # Fallback (if configured)
    if config.vision_fallback_url and config.vision_fallback_model:
        try:
            result = await _call_vision(config.vision_fallback_url, config.vision_fallback_model, image_b64, media_type, prompt)
            return f"[via fallback] {result}"
        except (httpx.ConnectError, httpx.ConnectTimeout):
            pass
        except Exception as e:
            return f"[Vision error: {e}]"

    return "[Vision unavailable — server unreachable.]"


async def describe_images(
    images: list[dict],
    prompt: str | None = None,
) -> list[dict]:
    """Describe multiple images via the vision model.

    Args:
        images: List of dicts with base64 and media_type keys
        prompt: Optional custom prompt

    Returns:
        List of dicts with added 'description' key
    """
    default_prompt = (
        "Describe this image in detail. If it's a diagram or flowchart, "
        "explain the structure, connections, and any text labels visible."
    )
    p = prompt or default_prompt

    results = []
    for img in images:
        if "error" in img:
            results.append({**img, "description": f"[Could not extract: {img['error']}]"})
            continue

        desc = await describe_image(img["base64"], img.get("media_type", "image/png"), p)
        results.append({**img, "description": desc})

    return results
