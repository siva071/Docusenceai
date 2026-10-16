"""Local Tesseract OCR for scanned PDF pages with page-aware error reporting."""

from __future__ import annotations

import os
import logging
from typing import Any, Callable

try:
    import pymupdf as fitz
except ImportError:  # Compatibility with earlier PyMuPDF releases.
    import fitz

OCRProgressCallback = Callable[[int, int, str], None]
DEFAULT_DPI = 240
DEFAULT_LANGUAGE = "eng"
DEFAULT_PAGE_SEGMENTATION_MODE = 6
LOGGER = logging.getLogger(__name__)


def check_ocr_runtime(language: str = DEFAULT_LANGUAGE) -> dict[str, object]:
    """Check local pytesseract, executable, and language availability."""
    try:
        import pytesseract
    except ImportError:
        return {
            "available": False,
            "message": "The Python OCR adapter is missing. Install project requirements with `python -m pip install -r requirements.txt`.",
        }
    try:
        import PIL.Image  # noqa: F401
    except ImportError:
        return {
            "available": False,
            "message": "The Python image dependency is missing. Install project requirements with `python -m pip install -r requirements.txt`.",
        }
    configured_executable = os.environ.get("TESSERACT_CMD")
    if configured_executable:
        pytesseract.pytesseract.tesseract_cmd = configured_executable
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        return {
            "available": False,
            "message": "Tesseract was not found. Install the Windows Tesseract application, add it to PATH, or set TESSERACT_CMD, then restart the app.",
        }
    except Exception:
        LOGGER.exception("Unable to query the local Tesseract runtime")
        return {
            "available": False,
            "message": "Tesseract could not be started. Check its installation and executable path.",
        }
    try:
        languages = pytesseract.get_languages(config="")
    except Exception:
        LOGGER.exception("Unable to enumerate Tesseract language data")
        return {
            "available": False,
            "message": "Tesseract is installed, but its language data could not be checked. Verify the tessdata installation and TESSDATA_PREFIX.",
        }
    if language not in languages:
        return {
            "available": False,
            "message": f"Tesseract language '{language}' is unavailable. Install its trained-data file and configure TESSDATA_PREFIX if needed.",
        }
    return {"available": True, "message": "Local Tesseract OCR is ready."}


def _image_coverage(page: Any) -> float:
    """Estimate the fraction of a page covered by embedded raster images."""
    page_area = float(page.rect.width * page.rect.height)
    if page_area <= 0:
        return 0.0
    image_area = 0.0
    seen_rects: set[tuple[float, float, float, float]] = set()
    for image in page.get_images(full=True):
        xref = image[0]
        for rectangle in page.get_image_rects(xref):
            rect = fitz.Rect(rectangle) & page.rect
            key = (round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y1, 2))
            if key not in seen_rects:
                seen_rects.add(key)
                image_area += max(0.0, rect.width * rect.height)
    return min(1.0, image_area / page_area)


def is_likely_scanned_page(
    selectable_text: str,
    image_coverage: float,
    min_selectable_characters: int = 20,
    min_image_coverage: float = 0.08,
) -> bool:
    """Return whether low-text content plus image coverage suggests a scan.

    A digital page with sufficient selectable text is not sent to OCR even when
    it includes embedded images (for example, a logo or chart).
    """
    character_count = len((selectable_text or "").strip())
    return character_count < min_selectable_characters and image_coverage >= min_image_coverage


def _word_records(data: dict[str, list[Any]], page_number: int, scale: float) -> list[dict[str, Any]]:
    """Convert pytesseract's TSV-style dictionaries to normalized word records."""
    words: list[dict[str, Any]] = []
    word_count = len(data.get("text", []))
    for index in range(word_count):
        text = str(data["text"][index]).strip()
        if not text:
            continue
        try:
            confidence = float(data.get("conf", ["-1"] * word_count)[index])
            left = float(data.get("left", [0] * word_count)[index]) * scale
            top = float(data.get("top", [0] * word_count)[index]) * scale
            width = float(data.get("width", [0] * word_count)[index]) * scale
            height = float(data.get("height", [0] * word_count)[index]) * scale
            block_number = int(data.get("block_num", [0] * word_count)[index])
            paragraph_number = int(data.get("par_num", [0] * word_count)[index])
            line_number = int(data.get("line_num", [0] * word_count)[index])
            word_number = int(data.get("word_num", [index + 1] * word_count)[index])
        except (TypeError, ValueError, IndexError):
            continue
        words.append(
            {
                "text": text,
                "confidence": confidence,
                "bbox": [left, top, left + width, top + height],
                "page_number": page_number,
                "block_number": block_number,
                "paragraph_number": paragraph_number,
                "line_number": line_number,
                "word_number": word_number,
            }
        )
    return words


def extract_text_with_ocr(
    pdf_bytes: bytes,
    page_numbers: list[int] | None = None,
    dpi: int = DEFAULT_DPI,
    language: str = DEFAULT_LANGUAGE,
    progress_callback: OCRProgressCallback | None = None,
    timeout_seconds: int = 60,
) -> dict[str, Any]:
    """OCR selected 1-based PDF page numbers locally with Tesseract.

    Only requested pages are rendered. A failure on one page is captured in
    that page's result and does not prevent OCR of later requested pages. The
    result includes positioned words for downstream layout/table inference.
    Tesseract itself must be installed separately and available on ``PATH`` or
    configured with ``pytesseract.pytesseract.tesseract_cmd``.
    """
    if dpi < 72 or dpi > 600:
        raise ValueError("dpi must be between 72 and 600")
    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:
        return {
            "pages": [],
            "warnings": ["OCR could not open this PDF. Check that the file is valid."],
            "fatal_error": str(exc),
        }

    selected = sorted(set(page_numbers or []))
    page_results: list[dict[str, Any]] = []
    warnings: list[str] = []
    try:
        if document.needs_pass:
            return {
                "pages": [],
                "warnings": ["OCR cannot process an encrypted PDF."],
                "fatal_error": "Encrypted PDFs are not supported.",
            }
        for page_number in selected:
            page_result: dict[str, Any] = {
                "page_number": page_number,
                "text": "",
                "words": [],
                "word_count": 0,
                "status": "ocr_failed",
                "error": None,
                "engine": "Tesseract",
                "language": language,
                "dpi": dpi,
                "page_width": None,
                "page_height": None,
            }
            try:
                if page_number < 1 or page_number > document.page_count:
                    raise ValueError(f"Page number {page_number} is outside this PDF.")
                page = document.load_page(page_number - 1)
                scale = dpi / 72.0
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                try:
                    from PIL import Image
                    import pytesseract
                    from pytesseract import Output
                except ImportError as exc:
                    raise RuntimeError(
                        "OCR dependencies are missing. Install project requirements including Pillow and pytesseract."
                    ) from exc

                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                configured_executable = os.environ.get("TESSERACT_CMD")
                if configured_executable:
                    pytesseract.pytesseract.tesseract_cmd = configured_executable
                data = pytesseract.image_to_data(
                    image,
                    lang=language,
                    config=f"--psm {DEFAULT_PAGE_SEGMENTATION_MODE} -c preserve_interword_spaces=1",
                    output_type=Output.DICT,
                    timeout=timeout_seconds,
                )
                words = _word_records(data, page_number, 72.0 / dpi)
                line_groups: dict[tuple[int, int, int], list[str]] = {}
                for word in words:
                    key = (
                        word["block_number"],
                        word["paragraph_number"],
                        word["line_number"],
                    )
                    line_groups.setdefault(key, []).append(word["text"])
                text = "\n".join(" ".join(line) for line in line_groups.values()).strip()
                if not text:
                    page_result["status"] = "ocr_no_text"
                    page_result["error"] = "Tesseract completed but found no readable text on this page."
                    warnings.append(f"Page {page_number}: OCR found no readable text; review this page.")
                else:
                    line_keys: dict[tuple[int, int, int], int] = {}
                    line_word_counts: dict[tuple[int, int, int], int] = {}
                    for word in words:
                        key = (
                            word["block_number"],
                            word["paragraph_number"],
                            word["line_number"],
                        )
                        if key not in line_keys:
                            line_keys[key] = len(line_keys) + 1
                            line_word_counts[key] = 0
                        line_word_counts[key] += 1
                        word["source_row"] = line_keys[key]
                        word["source_column"] = line_word_counts[key]
                    page_result.update(
                        {
                            "text": text,
                            "words": words,
                            "word_count": len(words),
                            "status": "ocr_text",
                            "page_width": float(page.rect.width),
                            "page_height": float(page.rect.height),
                        }
                    )
            except Exception as exc:
                LOGGER.exception("OCR failed on PDF page %s", page_number)
                error_message = _friendly_ocr_error(exc)
                page_result["error"] = error_message
                warnings.append(f"Page {page_number}: OCR failed - {error_message}")
            page_results.append(page_result)
            if progress_callback is not None:
                progress_callback(page_number, document.page_count, "ocr")
    finally:
        document.close()
    return {"pages": page_results, "warnings": warnings, "fatal_error": None}


def _friendly_ocr_error(error: Exception) -> str:
    """Convert common Tesseract/runtime errors into actionable messages."""
    try:
        import pytesseract
    except ImportError:
        pytesseract = None
    if pytesseract is not None and isinstance(error, pytesseract.TesseractNotFoundError):
        return (
            "Tesseract was not found. Install the Windows Tesseract OCR application, "
            "add its installation folder to PATH, or set the TESSERACT_CMD environment variable."
        )
    if pytesseract is not None and isinstance(error, pytesseract.TesseractError):
        message = str(error)
        if "Failed loading language" in message or "Error opening data file" in message:
            return (
                "The requested Tesseract language data is unavailable. Install the language pack "
                "and set TESSDATA_PREFIX if needed."
            )
        return f"Tesseract reported an OCR error: {message}"
    if "timeout" in str(error).casefold():
        return "OCR timed out on this page. Try a lower-resolution scan or retry the page."
    return "OCR failed on this page. Review the source image and check application logs for developer diagnostics."
