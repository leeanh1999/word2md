"""VietOCR without torch: the same model, read through onnxruntime.

VietOCR spells Vietnamese properly where Windows OCR cannot, but it runs on
torch, which unpacks to half a gigabyte - far too much to put in an executable
that downloads its own updates. `tools/export_vietocr_onnx.py` turns the model
into two small ONNX graphs, and this module runs them with onnxruntime (about
15 MB) and numpy, which the app already carries for pandas.

The reading is VietOCR's own, step for step: the line picture is scaled to 32
pixels high, the encoder turns it into a memory, and the decoder is asked for
one character at a time, always taking the likeliest, until it says the line
is over.
"""

from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import numpy as np

# Where the exported model lives: beside this module in a source checkout, and
# inside the unpacked bundle in the executable (build.py adds it there).
MODEL_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])) / (
    "src/models/vietocr"
)

# VietOCR's own preprocessing and decoding constants (vgg_seq2seq config).
_HEIGHT = 32
_MIN_WIDTH = 32
_MAX_WIDTH = 512
_ROUND_TO = 10
_MAX_LENGTH = 128
_SOS = 1
_EOS = 2
_FIRST_CHAR = 4


def available(model_dir: Path = MODEL_DIR) -> bool:
    """Whether the runtime and the model files are both here."""
    if not all((model_dir / name).is_file() for name in _FILES):
        return False
    try:
        import onnxruntime  # noqa: F401
    except Exception:  # noqa: BLE001 - a broken runtime is no runtime
        return False
    return True


_FILES = ("encoder.onnx", "decoder.onnx", "vocab.json")


class Recogniser:
    """Reads one line of text from a picture, the way vietocr's Predictor does."""

    def __init__(self, model_dir: Path = MODEL_DIR):
        import onnxruntime as ort

        options = ort.SessionOptions()
        options.log_severity_level = 3
        providers = ["CPUExecutionProvider"]
        self._encoder = ort.InferenceSession(
            str(model_dir / "encoder.onnx"), options, providers=providers
        )
        self._decoder = ort.InferenceSession(
            str(model_dir / "decoder.onnx"), options, providers=providers
        )
        vocab = json.loads((model_dir / "vocab.json").read_text(encoding="utf-8"))
        self._chars: str = vocab["chars"]

    @staticmethod
    def _prepare(image) -> np.ndarray:
        from PIL import Image

        image = image.convert("RGB")
        width, height = image.size
        new_width = int(_HEIGHT * float(width) / float(height))
        new_width = math.ceil(new_width / _ROUND_TO) * _ROUND_TO
        new_width = min(max(new_width, _MIN_WIDTH), _MAX_WIDTH)
        image = image.resize((new_width, _HEIGHT), Image.LANCZOS)
        pixels = np.asarray(image, dtype=np.float32).transpose(2, 0, 1) / 255
        return pixels[np.newaxis, ...]

    def predict(self, image) -> str:
        memory, hidden = self._encoder.run(None, {"image": self._prepare(image)})
        token = np.array([_SOS], dtype=np.int64)
        ids: list[int] = []
        for _ in range(_MAX_LENGTH + 1):
            logits, hidden = self._decoder.run(
                None, {"token": token, "hidden": hidden, "memory": memory}
            )
            best = int(logits[0].argmax())
            if best == _EOS:
                break
            ids.append(best)
            token[0] = best
        return "".join(
            self._chars[i - _FIRST_CHAR] for i in ids if i >= _FIRST_CHAR
        )
