"""Tests for VietOCR run through onnxruntime instead of torch."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import vietocr_onnx  # noqa: E402

# The real font only (msttcorefonts on Linux): VietOCR misreads look-alikes.
FONT = next((p for p in (Path("C:/Windows/Fonts/times.ttf"), Path("/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf"), Path("/usr/share/fonts/truetype/msttcorefonts/times.ttf")) if p.exists()), Path("C:/Windows/Fonts/times.ttf"))
LINES = [
    "Tài liệu này mô tả kết quả kiểm thử hệ thống",
    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
    "Số tiền: 25.500.000 đồng (Hai mươi lăm triệu)",
]


def line_picture(text: str, size: int = 50):
    from PIL import Image, ImageDraw, ImageFont

    font = ImageFont.truetype(str(FONT), size)
    left, top, right, bottom = font.getbbox(text)
    image = Image.new("RGB", (right - left + 24, bottom - top + 24), "white")
    ImageDraw.Draw(image).text((12 - left, 12 - top), text, font=font, fill="black")
    return image


class AvailabilityTests(unittest.TestCase):
    def test_a_folder_without_the_model_is_not_available(self):
        self.assertFalse(vietocr_onnx.available(Path(__file__).parent))


@unittest.skipUnless(vietocr_onnx.available(), "thiếu mô hình VietOCR (ONNX)")
class RecogniserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not FONT.exists():
            raise unittest.SkipTest("cần font Times New Roman")
        cls.recogniser = vietocr_onnx.Recogniser()

    def test_vietnamese_lines_are_read_with_their_accents(self):
        for text in LINES:
            with self.subTest(text=text):
                self.assertEqual(self.recogniser.predict(line_picture(text)), text)

    def test_a_tiny_crop_does_not_crash(self):
        from PIL import Image

        self.assertIsInstance(
            self.recogniser.predict(Image.new("RGB", (6, 6), "white")), str
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
