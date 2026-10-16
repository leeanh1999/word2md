"""Turn the VietOCR model into ONNX files the packaged app can read without torch.

    python tools/export_vietocr_onnx.py            # -> src/models/vietocr/

Run this on a development machine that has vietocr and torch installed (see
requirements-ocr-vi.txt) plus `onnx` and `onnxruntime`. It writes three files:

    encoder.onnx  the CNN and the GRU encoder: a line picture -> memory
    decoder.onnx  one step of the attention decoder: a token -> the next one
    vocab.json    the characters the token ids stand for

The decoder is exported a step at a time because the greedy loop around it is a
few lines of numpy (src/vietocr_onnx.py) and does not need to be in the graph.
Both graphs are then quantised to 8-bit weights, which takes them from 90 MB to
about a quarter of that; the export checks that the quantised model still reads
the sample lines the same way torch does before it keeps anything.
"""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

import numpy as np
import torch
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

MODEL = "vgg_seq2seq"
OUT = ROOT / "src" / "models" / "vietocr"
OPSET = 17

# Lines the export is checked against: accents, digits, punctuation, capitals.
SAMPLES = [
    "Tài liệu này mô tả kết quả kiểm thử hệ thống",
    "trong quý vừa qua, gồm 1.200.000 bản ghi.",
    "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
    "Độc lập - Tự do - Hạnh phúc",
    "Người đề nghị: Nguyễn Thị Hường",
    "Số tiền: 25.500.000 đồng (Hai mươi lăm triệu)",
]


class _Encoder(torch.nn.Module):
    def __init__(self, model):
        super().__init__()
        self.cnn = model.cnn
        self.encoder = model.transformer.encoder

    def forward(self, image):
        # vgg.forward, with the negative dims spelled out: the exporter writes
        # permute(-1, 0, 1) into the graph as is, and ONNX rejects it.
        vgg = self.cnn.model
        conv = vgg.last_conv_1x1(vgg.dropout(vgg.features(image)))
        conv = conv.transpose(3, 2).flatten(2).permute(2, 0, 1)
        outputs, hidden = self.encoder(conv)
        return outputs, hidden


class _DecoderStep(torch.nn.Module):
    """vietocr's Decoder.forward without the assert, which does not trace."""

    def __init__(self, model):
        super().__init__()
        self.decoder = model.transformer.decoder

    def forward(self, token, hidden, encoder_outputs):
        d = self.decoder
        embedded = d.embedding(token.unsqueeze(0))
        attention = d.attention(hidden, encoder_outputs).unsqueeze(1)
        weighted = torch.bmm(attention, encoder_outputs.permute(1, 0, 2)).permute(1, 0, 2)
        output, hidden = d.rnn(torch.cat((embedded, weighted), dim=2), hidden.unsqueeze(0))
        logits = d.fc_out(
            torch.cat((output.squeeze(0), weighted.squeeze(0), embedded.squeeze(0)), dim=1)
        )
        return logits, hidden.squeeze(0)


def _predictor():
    from vietocr.tool.config import Cfg
    from vietocr.tool.predictor import Predictor

    config = Cfg.load_config_from_name(MODEL)
    config["device"] = "cpu"
    config["cnn"]["pretrained"] = False
    return Predictor(config)


def _font(size: int):
    for name in ("times.ttf", "arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    raise SystemExit("No TrueType font with Vietnamese glyphs found to draw the samples.")


def sample_images() -> list[Image.Image]:
    """The sample lines drawn the way a 300 dpi scan of 12 pt text looks."""
    font = _font(50)
    images = []
    for text in SAMPLES:
        left, top, right, bottom = font.getbbox(text)
        image = Image.new("RGB", (right - left + 24, bottom - top + 24), "white")
        ImageDraw.Draw(image).text((12 - left, 12 - top), text, font=font, fill="black")
        images.append(image)
    return images


def export(out: Path, quantise: bool) -> None:
    import onnx
    from onnxruntime.quantization import QuantType, quantize_dynamic

    predictor = _predictor()
    model = predictor.model.eval()
    out.mkdir(parents=True, exist_ok=True)

    image = torch.rand(1, 3, 32, 160)
    with torch.no_grad():
        memory, hidden = _Encoder(model)(image)
    token = torch.tensor([1], dtype=torch.long)

    with tempfile.TemporaryDirectory() as scratch:
        scratch = Path(scratch)
        torch.onnx.export(
            _Encoder(model),
            (image,),
            scratch / "encoder.onnx",
            input_names=["image"],
            output_names=["memory", "hidden"],
            dynamic_axes={"image": {3: "width"}, "memory": {0: "steps"}},
            opset_version=OPSET,
            dynamo=False,
        )
        torch.onnx.export(
            _DecoderStep(model),
            (token, hidden, memory),
            scratch / "decoder.onnx",
            input_names=["token", "hidden", "memory"],
            output_names=["logits", "next_hidden"],
            dynamic_axes={"memory": {0: "steps"}},
            opset_version=OPSET,
            dynamo=False,
        )
        for name in ("encoder.onnx", "decoder.onnx"):
            if quantise:
                quantize_dynamic(
                    scratch / name,
                    out / name,
                    weight_type=QuantType.QUInt8,
                    # Shape inference loses the type after the GRU; it is float.
                    extra_options={"DefaultTensorType": onnx.TensorProto.FLOAT},
                )
            else:
                (out / name).write_bytes((scratch / name).read_bytes())

    (out / "vocab.json").write_text(
        json.dumps({"model": MODEL, "chars": predictor.vocab.chars}, ensure_ascii=False),
        encoding="utf-8",
    )

    # The point of the exercise: the files must read what torch reads.
    from src.vietocr_onnx import Recogniser

    onnx = Recogniser(out)
    worst = 0
    for picture, expected in zip(sample_images(), SAMPLES):
        by_torch = predictor.predict(picture)
        by_onnx = onnx.predict(picture)
        same = "==" if by_torch == by_onnx else "!="
        print(f"torch {by_torch!r}\nonnx  {by_onnx!r}  {same}  truth {expected!r}\n")
        worst += by_torch != by_onnx
    for name in ("encoder.onnx", "decoder.onnx", "vocab.json"):
        print(f"{name:14} {(out / name).stat().st_size / 1e6:6.1f} MB")
    if worst:
        print(f"{worst} of {len(SAMPLES)} samples read differently from torch.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("-o", "--output", type=Path, default=OUT)
    parser.add_argument(
        "--no-quantise", action="store_true", help="keep 32-bit weights (90 MB)"
    )
    args = parser.parse_args()
    export(args.output, quantise=not args.no_quantise)
    return 0


if __name__ == "__main__":
    sys.exit(main())
