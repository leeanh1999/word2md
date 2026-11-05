"""OCR for a PDF that is only a picture of a page.

A scanned PDF holds no text objects at all - pdfplumber comes back with nothing
- so the page has to be rasterised and read by an OCR engine. Reading it is two
jobs, and this module does them with two different tools:

    *where the words are* - Windows.Media.Ocr, the engine Windows ships with
        (the one behind PowerToys' Text Extractor), reached through the `winrt`
        bindings. It finds every word on the page and gives it a box, and it is
        very good at that whatever language the page is in.
    *what the words say*  - VietOCR. Windows has no Vietnamese recogniser, so
        it reads `Tài liệu` as `Tåi liéu`: the shape of the page survives and
        the letters do not. VietOCR is trained on Vietnamese, so each box is
        cropped out of the page picture and read again with it, and only the
        text is replaced. It runs as ONNX through onnxruntime (see
        `vietocr_onnx.py`), so the executable carries it without torch.

On Linux there is no Windows.Media.Ocr, so the word boxes come from Tesseract
instead (the `tesseract` command, its TSV output). With the `vie` traineddata
installed Tesseract already spells Vietnamese; VietOCR still re-reads each box
when its model is present.

Without the model files the engine's reading is kept as it is and the conversion
warns.
The boxes are scaled back to PDF points either way and handed to the same
layout pass that reads an ordinary PDF, so a scan comes out with the headings,
paragraphs, lists and tables a text PDF would have got - `pdf.py` does not care
which of the two it is reading.
"""

from __future__ import annotations


def _preload_recogniser() -> None:
    """Load onnxruntime before the winrt bindings get anywhere near it.

    The winrt package carries an old msvcp140.dll (14.29) and onnxruntime needs
    14.40 or later. Whichever loads first puts its copy in the process for
    good: import winrt first and onnxruntime dies on import with an access
    violation, while onnxruntime first brings in the newer runtime, which winrt
    is happy with. Nothing else in this module may import winrt above this line.
    """
    try:
        import onnxruntime  # noqa: F401
    except Exception:  # noqa: BLE001 - a broken runtime just means no VietOCR
        pass


_preload_recogniser()

import asyncio  # noqa: E402
import io  # noqa: E402
import re  # noqa: E402
import shutil  # noqa: E402
import subprocess  # noqa: E402
import sys  # noqa: E402
import threading  # noqa: E402
from concurrent.futures import ThreadPoolExecutor  # noqa: E402
from dataclasses import dataclass  # noqa: E402
from pathlib import Path  # noqa: E402

from . import vietocr_onnx  # noqa: E402
from .legacy import LegacyConversionError, external_env  # noqa: E402

# What the page is rasterised at. 300 dpi is what a scanner produces and what
# the engine reads best; below 200 the diacritics start to go.
OCR_DPI = 300
# The engine refuses an image bigger than this on a side.
_MAX_IMAGE_SIDE = 10000
# Preferred recognisers, best first, for a document in Vietnamese.
_PREFERRED = ("vi-VN", "vi", "en-US", "en")
# Which VietOCR model is exported to src/models/vietocr. The seq2seq one is a
# third the size of the transformer and read the samples here without a mistake.
VIETOCR_MODEL = "vgg_seq2seq"
# Points of margin around a crop, so no accent is cut off its letter.
_CROP_PADDING = 3
# A crop this thin or this short is punctuation, not a word worth re-reading.
_CROP_MINIMUM = 4
# A gap this many line heights wide ends one crop and starts the next. Whole
# table rows have to be split at them: VietOCR reads a line, and given a row
# the width of a page it answers with nonsense.
_CROP_GAP = 1.2
# What a list marker looks like. VietOCR is trained on prose and drops one from
# the front of a line, so it is put back from the Windows reading.
_MARKER = re.compile(r"^([•●▪◦‣⁃*+]|[-–—]|\d{1,3}[.)])$")
# Tesseract traineddata, best first, for a document in Vietnamese.
_TESSERACT_PREFERRED = ("vie", "eng")
# Seconds Tesseract may spend on one page before the page is given up.
_TESSERACT_TIMEOUT = 300
# Tesseract words below this confidence (0-100) are noise, not text.
_TESSERACT_MIN_CONFIDENCE = 10


class OcrError(LegacyConversionError):
    """This machine cannot OCR the page that was asked for."""


@dataclass
class OcrWord:
    """One recognised word, in PDF points with the origin at the top left."""

    text: str
    x0: float
    x1: float
    top: float
    bottom: float

    @property
    def height(self) -> float:
        return max(1.0, self.bottom - self.top)


# ------------------------------------------------------- a thread of our own

_worker: ThreadPoolExecutor | None = None
_worker_lock = threading.Lock()


def _run(work):
    """Run `work` on the one thread this module ever uses, and wait for it.

    Two things make a shared thread dangerous here. The OCR bindings want a
    multi-threaded COM apartment, while every thread that has driven Word or
    Excel through pywin32 is in a single-threaded one; and a thread that closes
    its apartment on the way out takes the cached OCR factories down with it,
    so the next call reaches into freed memory and the process dies. One thread
    of our own, kept alive for as long as the program runs, keeps the two apart
    and keeps the apartment standing.

    Never call this from inside the worker itself - it would wait on itself.
    """
    global _worker
    with _worker_lock:
        if _worker is None:
            _worker = ThreadPoolExecutor(
                max_workers=1, thread_name_prefix="word2md-ocr"
            )
    return _worker.submit(work).result()


# --------------------------------------------------------------- what we have


def _ocr_engine_class():
    from winrt.windows.media.ocr import OcrEngine

    return OcrEngine


def has_windows_ocr() -> bool:
    """Whether the Windows OCR engine can be reached from here."""
    try:
        _ocr_engine_class()
    except Exception:  # noqa: BLE001 - not Windows, or the bindings are absent
        return False
    return True


def find_tesseract() -> str | None:
    """The tesseract executable, or None."""
    return shutil.which("tesseract")


def has_tesseract() -> bool:
    return find_tesseract() is not None


def ocr_backend() -> str | None:
    """`windows`, `tesseract`, or None when nothing can find words here."""
    if has_windows_ocr():
        return "windows"
    if has_tesseract():
        return "tesseract"
    return None


def has_pypdfium2() -> bool:
    try:
        import pypdfium2  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def can_ocr() -> bool:
    """Both halves are needed: something to draw the page, something to read it."""
    return ocr_backend() is not None and has_pypdfium2() and bool(ocr_languages())


def _languages() -> list[str]:
    engine = _ocr_engine_class()
    return [
        str(language.language_tag) for language in engine.available_recognizer_languages
    ]


def _tesseract_languages() -> list[str]:
    tesseract = find_tesseract()
    if tesseract is None:
        return []
    try:
        done = subprocess.run(
            [tesseract, "--list-langs"],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
            env=external_env(),
        )
    except (OSError, subprocess.SubprocessError):
        return []
    # First line is a header ("List of available languages in ...").
    lines = (done.stdout or "").splitlines()[1:]
    return [line.strip() for line in lines if line.strip() and line.strip() != "osd"]


def ocr_languages() -> list[str]:
    """The recognisers installed on this machine.

    BCP-47 tags for Windows OCR (`vi-VN`), traineddata names for Tesseract (`vie`).
    """
    backend = ocr_backend()
    if backend == "tesseract":
        return _tesseract_languages()
    if backend is None:
        return []
    try:
        return _run(_languages)
    except Exception:  # noqa: BLE001 - no engine means no languages
        return []


def ocr_support_note() -> str:
    """One line telling the user what OCR looks like here."""
    if not has_windows_ocr() and sys.platform != "win32":
        return _tesseract_support_note()
    if not has_windows_ocr():
        return (
            "OCR: chưa dùng được (thiếu winrt - "
            "pip install -r requirements.txt trên Windows)."
        )
    if not has_pypdfium2():
        return "OCR: chưa dùng được (thiếu pypdfium2 để dựng ảnh trang)."
    languages = ocr_languages()
    if not languages:
        return "OCR: Windows chưa cài gói nhận dạng nào (Settings > Language)."
    if has_vietocr():
        return f"OCR: Windows dò chữ ({', '.join(languages)}), VietOCR đọc tiếng Việt."
    if not any(tag.lower().startswith("vi") for tag in languages):
        return (
            f"OCR: Windows nhận dạng được {', '.join(languages)}; chưa có tiếng Việt "
            "nên chữ có dấu sẽ sai (thiếu mô hình VietOCR trong src/models/vietocr)."
        )
    return f"OCR: Windows nhận dạng được {', '.join(languages)}."


def _tesseract_support_note() -> str:
    if not has_tesseract():
        return (
            "OCR: chưa dùng được (thiếu tesseract - "
            "sudo apt install tesseract-ocr tesseract-ocr-vie)."
        )
    if not has_pypdfium2():
        return "OCR: chưa dùng được (thiếu pypdfium2 để dựng ảnh trang)."
    languages = _tesseract_languages()
    if not languages:
        return "OCR: Tesseract chưa có gói ngôn ngữ nào (tesseract-ocr-eng/-vie)."
    if has_vietocr():
        return f"OCR: Tesseract dò chữ ({', '.join(languages)}), VietOCR đọc tiếng Việt."
    if "vie" not in languages:
        return (
            f"OCR: Tesseract nhận dạng được {', '.join(languages)}; chưa có tiếng Việt "
            "nên chữ có dấu sẽ sai (sudo apt install tesseract-ocr-vie)."
        )
    return f"OCR: Tesseract nhận dạng được {', '.join(languages)}."


def pick_language(languages: list[str]) -> str | None:
    """The best recogniser on hand for a Vietnamese document."""
    if languages and all("-" not in tag and len(tag) >= 3 for tag in languages):
        # Tesseract traineddata names (vie, eng, chi_sim ...), not BCP-47 tags.
        found = [tag for tag in _TESSERACT_PREFERRED if tag in languages]
        if found:
            return "+".join(found)
        return languages[0]
    lowered = {tag.lower(): tag for tag in languages}
    for wanted in _PREFERRED:
        for tag in lowered:
            if tag == wanted.lower() or tag.startswith(f"{wanted.lower()}-"):
                return lowered[tag]
    return languages[0] if languages else None


# ------------------------------------------------------------------- reading


def _render(document, index: int, dpi: int):
    """One page as a picture and as PNG bytes, with the pixels-per-point used."""
    page = document[index]
    scale = dpi / 72
    width, height = page.get_size()
    # Keep the picture inside what the engine will accept.
    longest = max(width, height) * scale
    if longest > _MAX_IMAGE_SIDE:
        scale *= _MAX_IMAGE_SIDE / longest
    image = page.render(scale=scale).to_pil()
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return image, buffer.getvalue(), scale


# ------------------------------------------------- reading the words again


def has_vietocr() -> bool:
    """Whether the Vietnamese recogniser - onnxruntime and the model - is here."""
    return vietocr_onnx.available()


_recogniser = None
_recogniser_failed = False


def _load_recogniser():
    """The VietOCR recogniser, loaded once and kept.

    Loading costs about a second, paid once per run and only by a document that
    actually has a scanned page in it.
    """
    global _recogniser, _recogniser_failed
    if _recogniser is not None or _recogniser_failed:
        return _recogniser
    if not has_vietocr():
        _recogniser_failed = True
        return None
    try:
        _recogniser = vietocr_onnx.Recogniser()
    except Exception:  # noqa: BLE001 - a broken model or runtime, no recogniser
        _recogniser_failed = True
        _recogniser = None
    return _recogniser


def _crops(row: list[OcrWord]) -> list[list[OcrWord]]:
    """Cut a line wherever a gap is too wide to be a space.

    What comes out is a cell of a table, or a whole line where there is no
    table. Either is something VietOCR can read; a full row of a statement is
    not.
    """
    heights = sorted(word.height for word in row)
    limit = heights[len(heights) // 2] * _CROP_GAP
    out: list[list[OcrWord]] = [[]]
    for word in row:
        if out[-1] and word.x0 - out[-1][-1].x1 > limit:
            out.append([])
        out[-1].append(word)
    return [group for group in out if group]


def _reread(group: list[OcrWord], image, scale: float, predictor) -> str | None:
    """Read one crop of the page again, in Vietnamese."""
    left = min(word.x0 for word in group) - _CROP_PADDING
    right = max(word.x1 for word in group) + _CROP_PADDING
    top = min(word.top for word in group) - _CROP_PADDING
    bottom = max(word.bottom for word in group) + _CROP_PADDING
    box = (
        max(0, int(left * scale)),
        max(0, int(top * scale)),
        min(image.width, int(right * scale)),
        min(image.height, int(bottom * scale)),
    )
    if box[2] - box[0] < _CROP_MINIMUM or box[3] - box[1] < _CROP_MINIMUM:
        return None
    try:
        text = str(predictor.predict(image.crop(box))).strip()
    except Exception:  # noqa: BLE001 - one unreadable crop, not one lost page
        return None
    return text or None


def _respell(group: list[OcrWord], text: str) -> list[OcrWord]:
    """Put a re-read crop back on the boxes it came from.

    Word for word where the two agree on how many there are - which keeps every
    box, and with it the columns and the type sizes the layout pass reads. Where
    they disagree the whole crop goes on one box spanning it, because a cell
    that lands in the right place matters more than the words inside it do.
    """
    marker = group[0].text
    if _MARKER.match(marker):
        # VietOCR is trained on prose: it swallows a leading bullet outright,
        # and runs a leading number into the word after it ("2.Phạm vi"). Both
        # cost the line its list item, so the marker is restored from the
        # Windows reading, which never had trouble with an ASCII dash or digit.
        rest = text[len(marker) :] if text.startswith(marker) else text
        text = f"{marker} {rest.lstrip()}"

    words = text.split()
    if len(words) == len(group):
        for word, spelling in zip(group, words):
            word.text = spelling
        return group

    first = group[0]
    first.text = text
    first.x0 = min(word.x0 for word in group)
    first.x1 = max(word.x1 for word in group)
    first.top = min(word.top for word in group)
    first.bottom = max(word.bottom for word in group)
    return [first]


def _rows(words: list[OcrWord]) -> list[list[OcrWord]]:
    """The words grouped into the lines they were printed on.

    By the middle of each box, never the top: a hyphen is a box three pixels
    tall sitting halfway down its line.
    """
    rows: list[list[OcrWord]] = []
    row: list[OcrWord] = []
    middle = 0.0
    for word in sorted(words, key=lambda item: ((item.top + item.bottom) / 2, item.x0)):
        centre = (word.top + word.bottom) / 2
        if row and centre - middle > row[0].height * 0.6:
            rows.append(sorted(row, key=lambda item: item.x0))
            row = []
        if not row:
            middle = centre
        row.append(word)
    if row:
        rows.append(sorted(row, key=lambda item: item.x0))
    return rows


def _refine(words: list[OcrWord], image, scale: float, predictor) -> list[OcrWord]:
    """Replace the Windows spelling of every word with VietOCR's reading."""
    refined: list[OcrWord] = []
    for row in _rows(words):
        for group in _crops(row):
            text = _reread(group, image, scale, predictor)
            refined.extend(_respell(group, text) if text else group)
    return refined


async def _software_bitmap(png: bytes):
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png)
    await writer.store_async()
    await writer.flush_async()
    writer.detach_stream()
    stream.seek(0)
    decoder = await BitmapDecoder.create_async(stream)
    return await decoder.get_software_bitmap_async()


def _engine(language: str | None):
    from winrt.windows.globalization import Language

    OcrEngine = _ocr_engine_class()
    if language:
        engine = OcrEngine.try_create_from_language(Language(language))
        if engine is not None:
            return engine
    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise OcrError(
            "Windows chưa cài gói nhận dạng chữ nào nên không OCR được "
            "(Settings > Time & language > Language)."
        )
    return engine


async def _read(engine, png: bytes, scale: float) -> list[OcrWord]:
    result = await engine.recognize_async(await _software_bitmap(png))
    words: list[OcrWord] = []
    for line in result.lines:
        for word in line.words:
            box = word.bounding_rect
            text = str(word.text).strip()
            if not text:
                continue
            words.append(
                OcrWord(
                    text=text,
                    x0=box.x / scale,
                    x1=(box.x + box.width) / scale,
                    top=box.y / scale,
                    bottom=(box.y + box.height) / scale,
                )
            )
    return words


async def _read_pages(
    source: Path, indexes: list[int], language: str | None, dpi: int, refine: bool
):
    import pypdfium2

    predictor = _load_recogniser() if refine else None
    document = pypdfium2.PdfDocument(str(source))
    try:
        engine = _engine(language)
        found: dict[int, list[OcrWord]] = {}
        for index in indexes:
            image, png, scale = _render(document, index, dpi)
            words = await _read(engine, png, scale)
            if predictor is not None and words:
                words = _refine(words, image, scale, predictor)
            found[index] = words
        return found
    finally:
        document.close()


def _tesseract_read(png: bytes, scale: float, language: str | None) -> list[OcrWord]:
    """One page through the tesseract CLI, as words in PDF points."""
    tesseract = find_tesseract()
    if tesseract is None:
        raise OcrError("Không tìm thấy tesseract (sudo apt install tesseract-ocr).")
    command = [tesseract, "stdin", "stdout", "--dpi", str(round(scale * 72))]
    if language:
        command += ["-l", language]
    command.append("tsv")
    try:
        done = subprocess.run(
            command,
            input=png,
            capture_output=True,
            timeout=_TESSERACT_TIMEOUT,
            check=False,
            env=external_env(),
        )
    except subprocess.TimeoutExpired as exc:
        raise OcrError("Tesseract đọc trang quá lâu, đã dừng.") from exc
    except OSError as exc:
        raise OcrError(f"Không chạy được tesseract: {exc}") from exc
    if done.returncode != 0:
        message = done.stderr.decode("utf-8", "replace").strip().splitlines()
        raise OcrError(f"Tesseract lỗi: {message[-1] if message else done.returncode}")
    return parse_tesseract_tsv(done.stdout.decode("utf-8", "replace"), scale)


def parse_tesseract_tsv(tsv: str, scale: float) -> list[OcrWord]:
    """Tesseract's TSV (level 5 = one word) turned into `OcrWord`s in points."""
    words: list[OcrWord] = []
    rows = tsv.splitlines()
    if not rows:
        return words
    header = rows[0].split("\t")
    try:
        col = {name: header.index(name) for name in (
            "level", "left", "top", "width", "height", "conf", "text"
        )}
    except ValueError:
        return words
    for row in rows[1:]:
        cells = row.split("\t")
        if len(cells) < len(header) or cells[col["level"]] != "5":
            continue
        text = cells[col["text"]].strip()
        try:
            confidence = float(cells[col["conf"]])
            left, top = int(cells[col["left"]]), int(cells[col["top"]])
            width, height = int(cells[col["width"]]), int(cells[col["height"]])
        except ValueError:
            continue
        if not text or confidence < _TESSERACT_MIN_CONFIDENCE:
            continue
        words.append(
            OcrWord(
                text=text,
                x0=left / scale,
                x1=(left + width) / scale,
                top=top / scale,
                bottom=(top + height) / scale,
            )
        )
    return words


def _tesseract_pages(
    source: Path, indexes: list[int], language: str | None, dpi: int, refine: bool
) -> dict[int, list[OcrWord]]:
    import pypdfium2

    predictor = _load_recogniser() if refine else None
    document = pypdfium2.PdfDocument(str(source))
    try:
        found: dict[int, list[OcrWord]] = {}
        for index in indexes:
            image, png, scale = _render(document, index, dpi)
            words = _tesseract_read(png, scale, language)
            if predictor is not None and words:
                words = _refine(words, image, scale, predictor)
            found[index] = words
        return found
    finally:
        document.close()


def ocr_pdf_pages(
    source: str | Path,
    indexes: list[int],
    language: str | None = None,
    dpi: int = OCR_DPI,
    refine: bool = True,
) -> dict[int, list[OcrWord]]:
    """Read the given pages (0-based) of a PDF as words placed on the page.

    One engine and one open document serve the whole batch: starting the engine
    again for every page costs more than reading the page does. With `refine`
    and VietOCR installed, every word is read a second time in Vietnamese and
    only its spelling changes.
    """
    if not indexes:
        return {}
    if not has_pypdfium2():
        raise OcrError("Không dựng được ảnh trang để OCR (thiếu pypdfium2).")
    backend = ocr_backend()
    if backend is None:
        if sys.platform == "win32":
            raise OcrError(
                "Không dùng được OCR của Windows (thiếu winrt; "
                "chạy pip install -r requirements.txt)."
            )
        raise OcrError(
            "Không có engine OCR (sudo apt install tesseract-ocr tesseract-ocr-vie)."
        )
    if language is None:
        language = pick_language(ocr_languages())
    if backend == "tesseract":
        try:
            return _tesseract_pages(Path(source), indexes, language, dpi, refine)
        except OcrError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise OcrError(f"OCR thất bại: {exc}") from exc
    try:
        return _run(
            lambda: asyncio.run(
                _read_pages(Path(source), indexes, language, dpi, refine)
            )
        )
    except OcrError:
        raise
    except Exception as exc:  # noqa: BLE001 - one failure, not one broken file
        raise OcrError(f"OCR thất bại: {exc}") from exc


__all__ = [
    "OCR_DPI",
    "VIETOCR_MODEL",
    "OcrError",
    "OcrWord",
    "can_ocr",
    "find_tesseract",
    "has_pypdfium2",
    "has_tesseract",
    "has_vietocr",
    "has_windows_ocr",
    "ocr_backend",
    "ocr_languages",
    "ocr_pdf_pages",
    "ocr_support_note",
    "parse_tesseract_tsv",
    "pick_language",
]
