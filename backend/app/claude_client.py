"""Calls Claude's vision-capable API to turn an uploaded drawing into a
structured BOM / cost-estimation payload (see schema.py)."""

import base64
import json
import mimetypes
import os
import re

import anthropic
from anthropic import Anthropic

MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5")
MAX_TOKENS = 8192

EXCEL_LIKE_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv"}
IMAGE_MIME_PREFIXES = ("image/",)

SYSTEM_PROMPT = """You are a senior quantity surveyor and production engineer at a \
joinery / woodworking factory that manufactures doors, cabinets, wardrobes and \
interior fit-out items. You are given an engineering drawing, shop drawing, door \
schedule, or specification sheet (as an image, PDF, or extracted spreadsheet text).

Your job: read everything relevant in the drawing and produce a complete Bill of \
Materials (BOM) and cost estimate as a single JSON object, with no prose before or \
after it and no markdown code fences.

Follow this exact JSON shape:

{
  "project": {
    "projectName": string,
    "documentRef": string,
    "client": string,
    "manufacturer": string,
    "date": string,               // best guess, "YYYY-MM-DD" or "" if unknown
    "currency": string,           // ISO-ish code, e.g. "SAR", "USD" — infer from context, default "SAR"
    "vatRate": number,            // fraction, e.g. 0.15 for 15%. Use the locally applicable rate if identifiable, else 0.15
    "notes": string
  },
  "items": [                      // top-level sellable items/sets (e.g. each door type, each cabinet type)
    {
      "itemNo": string,
      "drawingRef": string,
      "description": string,      // full spec paragraph: construction, core, finish, hardware summary
      "dimensions": string,       // key dimensions, e.g. "Wall: 1000x2200; Leaf: 914x2152x58"
      "unit": string,             // "Set", "Pcs", "No", etc.
      "qty": number,
      "unitCost": number          // your best-estimate unit sell price in `currency`; if the drawing states a price use it, otherwise estimate from typical market rates for the materials/construction described
    }
  ],
  "bom": [                        // detailed material breakdown per item (cutting list / parts list)
    {
      "itemNo": string,           // e.g. "1.01"
      "parentRef": string,        // which item.itemNo this part belongs to
      "assembly": string,         // e.g. "Leaf Assembly", "Frame Assembly", "Architraves"
      "partDescription": string,
      "materialSpec": string,
      "dimensions": string,
      "qtyPerUnit": number,
      "totalQty": number,         // qtyPerUnit * parent item's qty
      "unit": string,
      "notes": string
    }
  ],
  "hardware": [                   // ironmongery / accessories schedule, aggregated across items
    {
      "code": string,
      "category": string,
      "itemName": string,
      "brand": string,
      "material": string,
      "applicable": string,       // which item(s) this applies to
      "qtyPerUnit": number,
      "totalQty": number,
      "unit": string,
      "unitCost": number,         // your best-estimate cost in `currency`
      "scope": string             // short note on function/installation
    }
  ],
  "assumptions": [string]         // plainly list every guess you made: inferred prices, unclear dimensions,
                                  // ambiguous quantities, anything not explicitly on the drawing
}

Rules:
- Cover EVERYTHING you can identify in the drawing: every door/cabinet/item type, every material layer, every piece of hardware.
- Quantities in "bom" and "hardware" must be totals across the whole project (qtyPerUnit * relevant item qty), not per single unit, unless the field says otherwise.
- If the drawing gives explicit prices, use them and say so in "assumptions". If not, give a realistic estimate for the region/materials implied and clearly mark it as an estimate in "assumptions".
- If a section genuinely has nothing to report (e.g. no hardware visible), return an empty array for it — never omit a top-level key.
- Numbers must be plain JSON numbers, not strings, not formulas.
- Output strictly one JSON object. No markdown, no commentary, no trailing text.
"""

USER_INSTRUCTION = (
    "Analyze this drawing/document and return the BOM & cost estimation JSON "
    "described in your instructions. Extract every item, material and hardware "
    "component you can identify."
)


def _guess_media_type(filename: str, fallback: str) -> str:
    mime, _ = mimetypes.guess_type(filename)
    return mime or fallback


def _build_file_block(filename: str, content_type: str, data: bytes):
    """Returns an Anthropic content block for an image or PDF file."""
    media_type = content_type or _guess_media_type(filename, "application/octet-stream")
    b64 = base64.standard_b64encode(data).decode("ascii")

    if media_type == "application/pdf" or filename.lower().endswith(".pdf"):
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
        }

    if media_type.startswith(IMAGE_MIME_PREFIXES) or os.path.splitext(filename)[1].lower() in (
        ".png", ".jpg", ".jpeg", ".webp", ".gif",
    ):
        if media_type not in ("image/png", "image/jpeg", "image/webp", "image/gif"):
            media_type = "image/png"
        return {
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": b64},
        }

    return None


def _extract_json(text: str) -> dict:
    text = text.strip()
    # Strip ```json ... ``` fences defensively, in case the model adds them anyway.
    fence_match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Fall back to the outermost {...} block.
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate = text[start : end + 1]
        return json.loads(candidate)

    raise ValueError("Claude did not return parseable JSON")


def analyze_drawing(filename: str, content_type: str, data: bytes, extracted_text: str | None = None) -> dict:
    """Sends the uploaded drawing (or its extracted text, for spreadsheets) to
    Claude and returns the parsed structured BOM payload."""

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not configured on the server. Set it in the "
            "environment (see backend/.env.example) to enable drawing analysis."
        )

    client = Anthropic(api_key=api_key)

    user_content = []
    ext = os.path.splitext(filename)[1].lower()

    if extracted_text is not None or ext in EXCEL_LIKE_EXTENSIONS:
        text_payload = extracted_text or ""
        user_content.append(
            {
                "type": "text",
                "text": (
                    f"{USER_INSTRUCTION}\n\nThe source file is a spreadsheet named "
                    f"'{filename}'. Its extracted contents (sheet by sheet, row by row) "
                    f"are below:\n\n---\n{text_payload}\n---"
                ),
            }
        )
    else:
        file_block = _build_file_block(filename, content_type, data)
        if file_block is None:
            raise ValueError(f"Unsupported file type for '{filename}' ({content_type}).")
        user_content.append(file_block)
        user_content.append({"type": "text", "text": USER_INSTRUCTION})

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_content}],
        )
    except anthropic.AuthenticationError as exc:
        raise RuntimeError("Anthropic API key was rejected. Check ANTHROPIC_API_KEY.") from exc
    except anthropic.PermissionDeniedError as exc:
        raise RuntimeError("Anthropic API key does not have permission for this request.") from exc
    except anthropic.RateLimitError as exc:
        raise RuntimeError("Anthropic API rate limit reached. Try again shortly.") from exc
    except anthropic.APIStatusError as exc:
        message = getattr(exc, "message", None) or str(exc)
        if "credit balance" in message.lower():
            raise RuntimeError(
                "The configured Anthropic account has insufficient credits. "
                "Add credits in the Anthropic console, then retry."
            ) from exc
        raise RuntimeError(f"Anthropic API request failed: {message}") from exc
    except anthropic.APIConnectionError as exc:
        raise RuntimeError(f"Could not reach the Anthropic API: {exc}") from exc

    text_parts = [block.text for block in response.content if getattr(block, "type", None) == "text"]
    raw_text = "\n".join(text_parts)
    if not raw_text.strip():
        raise ValueError("Claude returned an empty response.")

    payload = _extract_json(raw_text)
    return payload
