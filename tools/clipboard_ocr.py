import sys
import platform
import subprocess
from io import BytesIO
import json

from PIL import Image, ImageGrab, ImageOps
import pytesseract
from pytesseract import Output, TesseractNotFoundError


TARGET = "экскаватор"


def get_clipboard_image():
    """Возвращает PIL.Image из буфера обмена (Win/macOS) или через xclip (Linux)."""
    try:
        data = ImageGrab.grabclipboard()
        if isinstance(data, Image.Image):
            return data
        if isinstance(data, list) and data:
            try:
                return Image.open(data[0])
            except Exception:
                pass
    except Exception:
        pass

    if platform.system().lower() == "linux":
        for mime in ("image/png", "image/jpeg", "image/bmp"):
            try:
                proc = subprocess.run(
                    ["xclip", "-selection", "clipboard", "-t", mime, "-o"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.DEVNULL,
                    check=True,
                )
                if proc.stdout:
                    return Image.open(BytesIO(proc.stdout))
            except Exception:
                continue

    return None


def preprocess_pil(img: Image.Image):
    """
    Возвращает предобработанное изображение + scale, чтобы вернуть координаты к оригиналу.
    """
    orig = img.convert("RGB")
    scale = 2
    w, h = orig.size
    up = orig.resize((w * scale, h * scale), Image.Resampling.LANCZOS)
    up = ImageOps.grayscale(up)
    up = ImageOps.autocontrast(up)
    return orig, up, scale


def _safe_float(x, default=-1.0):
    try:
        return float(x)
    except Exception:
        return default


def _union_bbox(boxes):
    # boxes: [(x1,y1,x2,y2), ...]
    x1 = min(b[0] for b in boxes)
    y1 = min(b[1] for b in boxes)
    x2 = max(b[2] for b in boxes)
    y2 = max(b[3] for b in boxes)
    return (x1, y1, x2, y2)


def find_all_target_fields(img_for_ocr: Image.Image, scale: int):
    """
    Ищет все "Экскаватор" и возвращает список:
    [
      {
        "label_text": ...,
        "label_conf": ...,
        "label_bbox": [x1,y1,x2,y2],
        "value_text": ...,
        "value_bbox": [x1,y1,x2,y2] | None
      }, ...
    ]
    """
    config = "--psm 6"
    d = pytesseract.image_to_data(
        img_for_ocr,
        lang="rus",
        config=config,
        output_type=Output.DICT
    )

    # Соберём токены
    tokens = []
    n = len(d["text"])
    for i in range(n):
        text = (d["text"][i] or "").strip()
        if not text:
            continue
        conf = _safe_float(d["conf"][i], -1.0)
        x = int(d["left"][i])
        y = int(d["top"][i])
        w = int(d["width"][i])
        h = int(d["height"][i])

        tokens.append({
            "i": i,
            "text": text,
            "text_l": text.lower(),
            "conf": conf,
            "x1": x,
            "y1": y,
            "x2": x + w,
            "y2": y + h,
            "page": int(d["page_num"][i]),
            "block": int(d["block_num"][i]),
            "par": int(d["par_num"][i]),
            "line": int(d["line_num"][i]),
            "word": int(d["word_num"][i]),
        })

    # Индексы-совпадения
    matches = [t for t in tokens if TARGET in t["text_l"]]

    results = []
    for m in matches:
        # label bbox в координатах оригинала
        label_bbox = [
            m["x1"] // scale,
            m["y1"] // scale,
            m["x2"] // scale,
            m["y2"] // scale,
        ]

        # Найдём "значение" справа в той же строке (те же page/block/par/line)
        same_line = [
            t for t in tokens
            if (t["page"], t["block"], t["par"], t["line"]) == (m["page"], m["block"], m["par"], m["line"])
        ]
        same_line.sort(key=lambda t: t["x1"])

        # Токены строго справа от метки
        right = [t for t in same_line if t["x1"] >= m["x2"] + 2]  # небольшой зазор

        # Если есть двоеточие/тире в следующем токене(ах), пропускаем их
        while right and right[0]["text"] in {":", "-", "—", "–"}:
            right.pop(0)

        # Берём "значение" как последовательность справа до большого разрыва
        value_tokens = []
        if right:
            prev_x2 = None
            for t in right:
                if prev_x2 is None:
                    value_tokens.append(t)
                    prev_x2 = t["x2"]
                    continue

                gap = t["x1"] - prev_x2
                # если разрыв слишком большой — считаем, что значение закончилось
                # (подстроено под скриншоты; при желании можно менять)
                if gap > 60:
                    break

                value_tokens.append(t)
                prev_x2 = t["x2"]

        if value_tokens:
            value_text = " ".join(t["text"] for t in value_tokens).strip()
            vb_pp = _union_bbox([(t["x1"], t["y1"], t["x2"], t["y2"]) for t in value_tokens])
            value_bbox = [vb_pp[0] // scale, vb_pp[1] // scale, vb_pp[2] // scale, vb_pp[3] // scale]
        else:
            value_text = ""
            value_bbox = None

        results.append({
            "label_text": m["text"],
            "label_conf": m["conf"],
            "label_bbox": label_bbox,
            "value_text": value_text,
            "value_bbox": value_bbox,
        })

    # Отсортируем сверху-вниз, слева-направо по bbox метки
    results.sort(key=lambda r: (r["label_bbox"][1], r["label_bbox"][0]))
    return results


def main():
    # Если у тебя иногда не видит PATH — раскомментируй и укажи путь:
    # pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    img = get_clipboard_image()
    if img is None:
        print("❌ В буфере обмена не найдено изображение.")
        sys.exit(1)

    _, img_pp, scale = preprocess_pil(img)

    try:
        results = find_all_target_fields(img_pp, scale=scale)
    except TesseractNotFoundError:
        print("❌ Python не нашёл tesseract.exe. Укажи путь явно в коде (см. комментарий).")
        sys.exit(2)

    if not results:
        print('Не найдено слово "Экскаватор".')
        sys.exit(3)

    # Человекочитаемый вывод
    for idx, r in enumerate(results, 1):
        print(f"[{idx}] label={r['label_text']!r} conf={r['label_conf']:.0f} bbox={tuple(r['label_bbox'])}")
        print(f"    value={r['value_text']!r} value_bbox={None if r['value_bbox'] is None else tuple(r['value_bbox'])}")

    # И машиночитаемый (удобно парсить)
    print("\nJSON:")
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
