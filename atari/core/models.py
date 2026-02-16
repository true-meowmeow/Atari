# Module: data models, actions, and OCR helpers.
# Main: Delay/Record/Action classes, action_from_dict, action_to_display.
# Example: from atari.core.models import Record, action_from_dict

import copy
import random
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Literal, Union

from PySide6.QtCore import QPoint, QRect
from PySide6.QtGui import QGuiApplication

from atari.core.geometry import rect_to_rel, rel_to_rect

DelayMode = Literal["fixed", "range"]

@dataclass
class Delay:
    mode: DelayMode = "fixed"
    a: float = 0.0
    b: float = 0.0  # used only for range

    def sample(self) -> float:
        if self.mode == "range":
            lo = min(self.a, self.b)
            hi = max(self.a, self.b)
            return random.uniform(lo, hi)
        return max(0.0, float(self.a))

    def to_dict(self) -> Dict[str, Any]:
        return {"mode": self.mode, "a": float(self.a), "b": float(self.b)}

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Delay":
        mode = d.get("mode", "fixed")
        a = float(d.get("a", 0.0))
        b = float(d.get("b", 0.0))
        if mode not in ("fixed", "range"):
            mode = "fixed"
        return Delay(mode=mode, a=a, b=b)


@dataclass
class RepeatSettings:
    enabled: bool = False
    count: int = 0  # 0 = infinite
    delay: Delay = field(default_factory=lambda: Delay("fixed", 0.5, 0.5))

    def to_dict(self) -> Dict[str, Any]:
        return {"enabled": bool(self.enabled), "count": int(self.count), "delay": self.delay.to_dict()}

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "RepeatSettings":
        return RepeatSettings(
            enabled=bool(d.get("enabled", False)),
            count=int(d.get("count", 0)),
            delay=Delay.from_dict(d.get("delay", {})),
        )


ActionType = Literal["base_area", "area", "key", "area_word", "wait", "wait_event"]

KeyKind = Literal["keys", "mouse"]
MouseButtonName = Literal["left", "middle", "right"]

# лучше держать это ДО dataclass'ов
DEFAULT_TRIGGER = {"kind": "mouse", "keys": [], "mouse_button": "left"}

_OCR_LANG_CACHE: Optional[List[str]] = None
_OCR_LANG_ALIASES = {
    "ru": "rus",
    "rus": "rus",
    "russian": "rus",
    "en": "eng",
    "eng": "eng",
    "english": "eng",
}


def _split_ocr_lang_spec(lang: str) -> List[str]:
    raw = str(lang or "").strip().lower()
    if not raw:
        return []
    out: List[str] = []
    seen = set()
    for part in re.split(r"[+,\s;|]+", raw):
        code = _OCR_LANG_ALIASES.get(part, part).strip().lower()
        if not code:
            continue
        if code in seen:
            continue
        seen.add(code)
        out.append(code)
    return out


def get_installed_ocr_languages(force_refresh: bool = False) -> List[str]:
    global _OCR_LANG_CACHE
    if _OCR_LANG_CACHE is not None and not force_refresh:
        return list(_OCR_LANG_CACHE)

    langs: List[str] = []
    try:
        import pytesseract
        from pytesseract import TesseractNotFoundError
    except Exception:
        _OCR_LANG_CACHE = []
        return []

    try:
        try:
            raw_langs = pytesseract.get_languages(config="")
        except TypeError:
            raw_langs = pytesseract.get_languages()
    except TesseractNotFoundError:
        raw_langs = []
    except Exception:
        raw_langs = []

    seen = set()
    for raw in (raw_langs or []):
        code = str(raw or "").strip().lower()
        if not code:
            continue
        if code in ("osd",):
            continue
        if code in seen:
            continue
        seen.add(code)
        langs.append(code)

    _OCR_LANG_CACHE = langs
    return list(_OCR_LANG_CACHE)


def is_ocr_available(force_refresh: bool = False) -> bool:
    return bool(get_installed_ocr_languages(force_refresh=force_refresh))


def normalize_ocr_lang_spec(lang: str, available: Optional[List[str]] = None) -> str:
    requested = _split_ocr_lang_spec(lang)

    if available is None:
        available = get_installed_ocr_languages(force_refresh=False)

    available_map: Dict[str, str] = {}
    for raw in (available or []):
        code = str(raw or "").strip()
        if not code:
            continue
        available_map[code.lower()] = code.lower()

    selected: List[str] = []
    for code in requested:
        if not available_map:
            if code not in selected:
                selected.append(code)
            continue
        if code in available_map and code not in selected:
            selected.append(code)

    if not selected:
        if available_map:
            if "rus" in available_map:
                selected = ["rus"]
            else:
                selected = [next(iter(available_map.keys()))]
        else:
            selected = ["rus"]

    return "+".join(selected)


def normalize_trigger(spec: Optional[dict]) -> dict:
    if not isinstance(spec, dict):
        spec = {}
    kind = spec.get("kind", "mouse")
    if kind not in ("mouse", "keys"):
        kind = "mouse"
    keys = [str(k) for k in (spec.get("keys", []) or [])][:10]
    mb = spec.get("mouse_button", None)
    if kind == "mouse":
        if mb not in ("left", "middle", "right"):
            mb = "left"
    else:
        mb = None
    return {"kind": kind, "keys": keys, "mouse_button": mb}


@dataclass
class BaseAreaAction:
    type: ActionType = "base_area"
    x1: int = 0;
    y1: int = 0;
    x2: int = 0;
    y2: int = 0

    click: bool = True  # <-- по умолчанию ЛКМ (как ты просил)
    trigger: Dict[str, Any] = field(default_factory=lambda: dict(DEFAULT_TRIGGER))

    def rect(self) -> QRect:
        return QRect(QPoint(self.x1, self.y1), QPoint(self.x2, self.y2)).normalized()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "type": "base_area",
            "x1": int(self.x1), "y1": int(self.y1), "x2": int(self.x2), "y2": int(self.y2),
            "click": bool(self.click),
            "trigger": normalize_trigger(self.trigger),
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "BaseAreaAction":
        return BaseAreaAction(
            x1=int(d.get("x1", 0)), y1=int(d.get("y1", 0)),
            x2=int(d.get("x2", 0)), y2=int(d.get("y2", 0)),
            click=bool(d.get("click", True)),
            trigger=normalize_trigger(d.get("trigger", DEFAULT_TRIGGER)),
        )


@dataclass
class WordAreaAction:
    type: ActionType = "area_word"
    x1: int = 0;
    y1: int = 0;
    x2: int = 0;
    y2: int = 0
    word: str = ""
    index: int = 1
    count: int = 1

    coord: Literal["abs", "rel"] = "rel"
    rx1: float = 0.0;
    ry1: float = 0.0;
    rx2: float = 1.0;
    ry2: float = 1.0

    click: bool = False
    multiplier: int = 1
    delay: Delay = field(default_factory=lambda: Delay("fixed", 0.1, 0.1))
    button: MouseButtonName = "left"  # оставим для старых json
    trigger: Dict[str, Any] = field(default_factory=lambda: dict(DEFAULT_TRIGGER))  # <-- НОВОЕ

    ocr_lang: str = "rus"
    search_infinite: bool = True
    search_max_tries: int = 100
    search_on_fail: Literal["retry", "error", "action"] = "retry"
    on_fail_actions: List[Dict[str, Any]] = field(default_factory=list)
    on_fail_post_mode: Literal["none", "stop", "repeat"] = "none"

    def search_rect_abs(self) -> QRect:
        return QRect(QPoint(self.x1, self.y1), QPoint(self.x2, self.y2)).normalized()

    def search_rect_global(self, base: Optional[QRect]) -> QRect:
        if self.coord == "rel" and base and base.isValid() and base.width() >= 2 and base.height() >= 2:
            return rel_to_rect(base, self.rx1, self.ry1, self.rx2, self.ry2)
        return self.search_rect_abs()

    def to_dict(self) -> Dict[str, Any]:
        trig = normalize_trigger(self.trigger)

        # совместимость со старым "button"
        btn = self.button
        if trig["kind"] == "mouse" and trig["mouse_button"] in ("left", "middle", "right"):
            btn = trig["mouse_button"]

        out = {
            "type": "area_word",
            "word": str(self.word),
            "index": int(self.index),
            "count": int(self.count),
            "click": bool(self.click),
            "multiplier": int(self.multiplier),
            "delay": self.delay.to_dict(),
            "button": btn,
            "trigger": trig,
            "ocr_lang": str(self.ocr_lang),
            "search_infinite": bool(self.search_infinite),
            "search_max_tries": max(1, int(self.search_max_tries)),
            "search_on_fail": self.search_on_fail if self.search_on_fail in ("retry", "error", "action") else "retry",
            "on_fail_actions": [copy.deepcopy(x) for x in (self.on_fail_actions or []) if isinstance(x, dict)],
            "on_fail_post_mode": (
                self.on_fail_post_mode if self.on_fail_post_mode in ("none", "stop", "repeat") else "none"
            ),
        }

        if self.coord == "rel":
            out.update({
                "coord": "rel",
                "rx1": float(self.rx1), "ry1": float(self.ry1),
                "rx2": float(self.rx2), "ry2": float(self.ry2),
            })
        else:
            out.update({
                "coord": "abs",
                "x1": int(self.x1), "y1": int(self.y1),
                "x2": int(self.x2), "y2": int(self.y2),
            })
        return out

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "WordAreaAction":
        btn = d.get("button", "left")
        if btn not in ("left", "middle", "right"):
            btn = "left"

        trig_raw = d.get("trigger", None)
        if isinstance(trig_raw, dict):
            trig = normalize_trigger(trig_raw)
        else:
            trig = normalize_trigger({"kind": "mouse", "keys": [], "mouse_button": btn})

        if "delay" in d:
            delay = Delay.from_dict(d.get("delay", {}))
        else:
            delay = Delay("fixed", 0.1, 0.1)

        idx = int(d.get("index", 1))
        if idx < 1:
            idx = 1

        count_raw = d.get("count", None)
        if count_raw is None:
            count = 0
        else:
            try:
                count = int(count_raw)
            except Exception:
                count = 1
            if count < 1:
                count = 0

        try:
            search_max_tries = int(d.get("search_max_tries", 100))
        except Exception:
            search_max_tries = 100
        if search_max_tries < 1:
            search_max_tries = 1
        search_on_fail = str(d.get("search_on_fail", "retry") or "retry")
        if search_on_fail not in ("retry", "error", "action"):
            search_on_fail = "retry"
        ofa_raw = d.get("on_fail_actions", [])
        if isinstance(ofa_raw, list):
            on_fail_actions = [copy.deepcopy(x) for x in ofa_raw if isinstance(x, dict)]
        else:
            on_fail_actions = []
        on_fail_post_mode = str(d.get("on_fail_post_mode", "none") or "none")
        if on_fail_post_mode not in ("none", "stop", "repeat"):
            on_fail_post_mode = "none"

        coord = d.get("coord", None)
        is_rel = (coord == "rel") or any(k in d for k in ("rx1", "ry1", "rx2", "ry2"))

        obj = WordAreaAction(
            x1=int(d.get("x1", 0)), y1=int(d.get("y1", 0)),
            x2=int(d.get("x2", 0)), y2=int(d.get("y2", 0)),
            word=str(d.get("word", "") or ""),
            index=idx,
            count=count,
            click=bool(d.get("click", False)),
            multiplier=max(1, int(d.get("multiplier", 1))),
            delay=delay,
            button=btn,
            trigger=trig,
            ocr_lang=str(d.get("ocr_lang", "rus") or "rus"),
            search_infinite=bool(d.get("search_infinite", True)),
            search_max_tries=search_max_tries,
            search_on_fail=search_on_fail,
            on_fail_actions=on_fail_actions,
            on_fail_post_mode=on_fail_post_mode,
        )

        if is_rel:
            obj.coord = "rel"
            obj.rx1 = float(d.get("rx1", 0.0));
            obj.ry1 = float(d.get("ry1", 0.0))
            obj.rx2 = float(d.get("rx2", 1.0));
            obj.ry2 = float(d.get("ry2", 1.0))
        else:
            obj.coord = "abs"

        return obj

    @staticmethod
    def _norm(s: str) -> str:
        s = (s or "").strip()
        s = unicodedata.normalize("NFKC", s)  # нормализуем юникод
        s = s.casefold()  # регистронезависимо (лучше чем lower)
        s = s.replace("ё", "е")  # опционально, но часто полезно для RU OCR
        # выкидываем всё кроме букв/цифр/подчёркивания (юникод-ок)
        return re.sub(r"[^\w]+", "", s, flags=re.UNICODE)

    def resolve_target_rect_global(
            self,
            base: Optional[QRect] = None,
            debug: bool = False,
            debug_tag: str = "",
            dpr_override: Optional[float] = None,  # NEW
    ) -> Optional[QRect]:

        rect = self.search_rect_global(base)

        if not rect.isValid() or rect.width() < 5 or rect.height() < 5:
            return None

        w_raw = (self.word or "").strip()
        if not w_raw:
            return None

        available_langs = get_installed_ocr_languages(force_refresh=False)
        if not available_langs:
            return None
        lang_for_ocr = normalize_ocr_lang_spec(getattr(self, "ocr_lang", "rus"), available=available_langs)

        try:
            from PIL import Image, ImageGrab, ImageOps, ImageFilter, ImageStat
            import pytesseract
            from pytesseract import Output, TesseractNotFoundError
            from difflib import SequenceMatcher
        except Exception:
            return None

        # --- HiDPI: Qt coords (DIP) -> физические пиксели для ImageGrab ---
        dpr = 1.0
        if dpr_override is not None:
            try:
                dpr = float(dpr_override)
                if dpr <= 0:
                    dpr = 1.0
            except Exception:
                dpr = 1.0
        else:
            # fallback (если не передали dpr_override)
            try:
                scr = QGuiApplication.screenAt(rect.center())
                if scr:
                    dpr = float(scr.devicePixelRatio())
            except Exception:
                dpr = 1.0

        # bbox для Pillow: left, top, right(exclusive), bottom(exclusive) в ФИЗ. пикселях
        bbox = (
            int(round(rect.left() * dpr)),
            int(round(rect.top() * dpr)),
            int(round((rect.right() + 1) * dpr)),
            int(round((rect.bottom() + 1) * dpr)),
        )

        # --- grab screenshot ---
        try:
            try:
                img = ImageGrab.grab(bbox=bbox, all_screens=True)
            except TypeError:
                img = ImageGrab.grab(bbox=bbox)
        except Exception:
            return None

        # --- helper: Otsu threshold (без numpy) ---
        def otsu_threshold(gray_img) -> int:
            hist = gray_img.histogram()  # 256 bins
            total = sum(hist)
            if total <= 0:
                return 128

            sum_all = 0
            for i, h in enumerate(hist):
                sum_all += i * h

            sum_b = 0
            w_b = 0
            var_max = -1.0
            thr = 128

            for t in range(256):
                w_b += hist[t]
                if w_b == 0:
                    continue
                w_f = total - w_b
                if w_f == 0:
                    break

                sum_b += t * hist[t]
                m_b = sum_b / w_b
                m_f = (sum_all - sum_b) / w_f

                var_between = w_b * w_f * (m_b - m_f) * (m_b - m_f)
                if var_between > var_max:
                    var_max = var_between
                    thr = t
            return int(thr)

        # --- adaptive upscale (важно для маленьких зон/шрифтов) ---
        try:
            orig = img.convert("RGB")
            ow, oh = orig.size
            min_dim = min(ow, oh)

            # Чем меньше зона — тем сильнее апскейл
            if min_dim < 220:
                scale = 4
            elif min_dim < 420:
                scale = 3
            elif min_dim < 900:
                scale = 2
            else:
                scale = 1

            # Pillow compatibility
            resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)

            up = orig.resize((max(1, ow * scale), max(1, oh * scale)), resampling)
            up = ImageOps.grayscale(up)
            up = ImageOps.autocontrast(up)

            # если фон тёмный, текст светлый — инвертируем
            try:
                mean = ImageStat.Stat(up).mean[0]
                if mean < 110:
                    up = ImageOps.invert(up)
            except Exception:
                pass

            # чуть “поджать” резкость для тонких букв
            up = up.filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=2))

            # бинаризация (Otsu)
            thr = otsu_threshold(up)
            up = up.point(lambda p: 255 if p > thr else 0)

        except Exception:
            return None

        target_norm = self._norm(w_raw)
        if not target_norm:
            return None
        try:
            expected_count = int(getattr(self, "count", 0))
        except Exception:
            expected_count = 0
        if expected_count < 0:
            expected_count = 0

        def safe_float(x, default=-1.0):
            try:
                return float(x)
            except Exception:
                return default

        # --- OCR: пробуем несколько psm (на мелком тексте реально помогает) ---
        best_matches: List[Tuple[int, int, int, int, float]] = []  # y1,x1,x2,y2,score
        psm_list = (6, 7, 11)  # 6=block, 7=single line, 11=sparse text

        for psm in psm_list:
            try:
                data = pytesseract.image_to_data(
                    up,
                    lang=lang_for_ocr,
                    config=f"--oem 3 --psm {psm} -c preserve_interword_spaces=1",
                    output_type=Output.DICT
                )
                if debug:
                    raw_tokens = []
                    n = len(data.get("text", []))
                    for i in range(n):
                        t = (data["text"][i] or "").strip()
                        if t:
                            raw_tokens.append(t)

                    raw_join = re.sub(r"\s+", " ", " ".join(raw_tokens)).strip()
                    if len(raw_join) > 400:
                        raw_join = raw_join[:400] + "…"

                    norm_tokens = []
                    for t in raw_tokens:
                        nt = self._norm(t)
                        if nt:
                            norm_tokens.append(nt)
                    norm_join = " ".join(norm_tokens)
                    if len(norm_join) > 400:
                        norm_join = norm_join[:400] + "…"

                    print(f"{debug_tag}psm={psm} OCR(raw)='{raw_join}'", flush=True)
                    print(f"{debug_tag}psm={psm} OCR(norm)='{norm_join}'", flush=True)

            except TesseractNotFoundError:
                return None
            except Exception:
                continue

            # ---------- NEW: надёжный матчинг слова/склейки слов ----------
            def _get_int_field(key: str, i: int, default: int = 0) -> int:
                try:
                    arr = data.get(key, None)
                    if not arr:
                        return default
                    return int(arr[i])
                except Exception:
                    return default

            def _gap_ok(prev_box, cur_box) -> bool:
                # prev_box/cur_box = (x1,y1,x2,y2,h)
                # допускаем небольшой разрыв между словами (пробел)
                px1, py1, px2, py2, ph = prev_box
                cx1, cy1, cx2, cy2, ch = cur_box
                gap = cx1 - px2
                # если слова налезают/почти касаются — ок
                if gap <= 2:
                    return True
                # иначе допускаем разрыв не больше ~1.4 высоты строки
                return gap <= int(max(ph, ch) * 1.4)

            def _threshold_for_target(tlen: int) -> float:
                # пороги под одно слово: чем длиннее цель — тем чуть мягче по ratio
                if tlen >= 12:
                    return 0.82
                if tlen >= 8:
                    return 0.85
                return 0.90  # короткие слова должны совпадать точнее

            tlen = len(target_norm)
            THR = _threshold_for_target(tlen)

            # минимальная длина кандидата (режем одиночные буквы и мусор)
            MIN_LEN = max(4, int(tlen * 0.60))  # для "Возрождение"(11) -> >=6
            MAX_LEN = int(tlen * 1.40) + 2  # чтобы не уходить в супер-длинные склейки
            MAX_JOIN = 3  # склеиваем до 3 соседних OCR-токенов

            # собираем слова OCR с координатами + принадлежностью строке
            words = []
            n = len(data.get("text", []))
            for i in range(n):
                txt = (data["text"][i] or "").strip()
                if not txt:
                    continue

                conf = safe_float(data.get("conf", ["-1"])[i], -1.0)
                if conf >= 0 and conf < 15:
                    continue

                x = int(data["left"][i])
                y = int(data["top"][i])
                ww = int(data["width"][i])
                hh = int(data["height"][i])

                norm = self._norm(txt)
                if not norm:
                    continue

                line_key = (
                    _get_int_field("block_num", i, 0),
                    _get_int_field("par_num", i, 0),
                    _get_int_field("line_num", i, 0),
                )

                words.append({
                    "i": i,
                    "raw": txt,
                    "norm": norm,
                    "conf": conf,
                    "x1": x, "y1": y, "x2": x + ww, "y2": y + hh,
                    "h": hh,
                    "line": line_key,
                })

            # ничего похожего не распознали
            tokens = []
            if words:
                # сортируем визуально (строка -> слева направо)
                words.sort(key=lambda w: (w["line"][0], w["line"][1], w["line"][2], w["y1"], w["x1"]))

                # группируем по строкам
                by_line = {}
                for w in words:
                    by_line.setdefault(w["line"], []).append(w)

                for _lk, ws in by_line.items():
                    m = len(ws)
                    for a0 in range(m):
                        cand_norm = ""
                        cand_raw_parts = []
                        confs = []
                        x1 = y1 = 10 ** 9
                        x2 = y2 = -10 ** 9

                        prev_box = None

                        for k in range(MAX_JOIN):
                            j = a0 + k
                            if j >= m:
                                break

                            w = ws[j]
                            cur_box = (w["x1"], w["y1"], w["x2"], w["y2"], w["h"])

                            # если это не первый токен склейки — проверяем, что он "рядом", а не через пол-экрана
                            if prev_box is not None and not _gap_ok(prev_box, cur_box):
                                break
                            prev_box = cur_box

                            cand_norm += w["norm"]
                            cand_raw_parts.append(w["raw"])
                            confs.append(max(0.0, w["conf"]))

                            if len(cand_norm) > MAX_LEN:
                                break

                            # быстрые фильтры длины
                            if len(cand_norm) < MIN_LEN:
                                continue

                            # длина кандидата должна быть близка к цели
                            if abs(len(cand_norm) - tlen) > max(2, int(tlen * 0.30)):
                                continue

                            ratio = SequenceMatcher(None, cand_norm, target_norm).ratio()
                            if ratio < THR:
                                continue

                            # bbox объединённый
                            x1 = min(x1, w["x1"])
                            y1 = min(y1, w["y1"])
                            x2 = max(x2, w["x2"])
                            y2 = max(y2, w["y2"])

                            avg_conf = (sum(confs) / max(1, len(confs))) if confs else 0.0
                            score = (avg_conf / 100.0) * 0.6 + ratio * 0.4

                            if debug:
                                cand_raw = " ".join(cand_raw_parts)
                                print(
                                    f"{debug_tag}HIT cand='{cand_raw}' norm='{cand_norm}' "
                                    f"len={len(cand_norm)} ratio={ratio:.2f} conf={avg_conf:.1f} join={k + 1}",
                                    flush=True
                                )

                            tokens.append((x1, y1, x2, y2, score))

            # ---------- /NEW ----------

            if not tokens:
                continue

            # перевод координат: up(px) -> orig(phys px) -> Qt(DIP)
            matches = []
            for x1, y1, x2, y2, score in tokens:
                # из апскейла обратно в физ. пиксели
                lx1_phys = x1 / scale
                ly1_phys = y1 / scale
                lx2_phys = x2 / scale
                ly2_phys = y2 / scale

                # из физ. пикселей обратно в DIP
                lx1 = lx1_phys / dpr
                ly1 = ly1_phys / dpr
                lx2 = lx2_phys / dpr
                ly2 = ly2_phys / dpr

                gx1 = int(round(rect.left() + lx1))
                gy1 = int(round(rect.top() + ly1))
                gx2 = int(round(rect.left() + lx2))
                gy2 = int(round(rect.top() + ly2))

                matches.append((gy1, gx1, gx2, gy2, score))

            if matches:
                if expected_count > 0 and len(matches) < expected_count:
                    if debug:
                        print(
                            f"{debug_tag}matches={len(matches)} expected={expected_count} -> retry",
                            flush=True
                        )
                    continue
                # берём лучший набор по суммарному score
                matches.sort(key=lambda b: (-b[4], b[0], b[1]))
                best_matches = matches
                break  # первый удачный psm обычно уже ок

        if not best_matches:
            return None

        # сорт “визуально”: сверху-вниз/слева-направо
        best_matches.sort(key=lambda b: (b[0], b[1]))

        idx0 = max(0, int(self.index) - 1)
        if idx0 >= len(best_matches):
            idx0 = len(best_matches) - 1

        y1, x1, x2, y2, _score = best_matches[idx0]
        return QRect(QPoint(x1, y1), QPoint(x2, y2)).normalized()


@dataclass
class AreaAction:
    type: ActionType = "area"

    # legacy absolute (для старых json)
    x1: int = 0;
    y1: int = 0;
    x2: int = 0;
    y2: int = 0

    # new relative (0..1 внутри base_area)
    coord: Literal["abs", "rel"] = "rel"
    rx1: float = 0.0;
    ry1: float = 0.0;
    rx2: float = 1.0;
    ry2: float = 1.0

    click: bool = False
    multiplier: int = 1
    delay: Delay = field(default_factory=lambda: Delay("fixed", 0.1, 0.1))
    trigger: Dict[str, Any] = field(default_factory=lambda: dict(DEFAULT_TRIGGER))

    def rect_abs(self) -> QRect:
        return QRect(QPoint(self.x1, self.y1), QPoint(self.x2, self.y2)).normalized()

    def rect_global(self, base: Optional[QRect]) -> QRect:
        if self.coord == "rel" and base and base.isValid() and base.width() >= 2 and base.height() >= 2:
            return rel_to_rect(base, self.rx1, self.ry1, self.rx2, self.ry2)
        return self.rect_abs()

    @staticmethod
    def from_global(rect: QRect, base: QRect, click: bool = False, trigger: Optional[dict] = None,
                    multiplier: int = 1, delay: Optional[Delay] = None) -> "AreaAction":
        rx1, ry1, rx2, ry2 = rect_to_rel(base, rect)
        return AreaAction(
            coord="rel", rx1=rx1, ry1=ry1, rx2=rx2, ry2=ry2,
            click=bool(click),
            multiplier=max(1, int(multiplier)),
            delay=Delay.from_dict(delay.to_dict()) if delay else Delay("fixed", 0.1, 0.1),
            trigger=normalize_trigger(trigger or DEFAULT_TRIGGER),
        )

    def to_dict(self) -> Dict[str, Any]:
        out = {
            "type": "area",
            "click": bool(self.click),
            "multiplier": int(self.multiplier),
            "delay": self.delay.to_dict(),
            "trigger": normalize_trigger(self.trigger),
        }
        if self.coord == "rel":
            out.update({"coord": "rel", "rx1": float(self.rx1), "ry1": float(self.ry1), "rx2": float(self.rx2),
                        "ry2": float(self.ry2)})
        else:
            out.update({"coord": "abs", "x1": int(self.x1), "y1": int(self.y1), "x2": int(self.x2), "y2": int(self.y2)})
        return out

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "AreaAction":
        if "delay" in d:
            delay = Delay.from_dict(d.get("delay", {}))
        else:
            delay = Delay("fixed", 0.1, 0.1)

        coord = d.get("coord", None)
        if coord == "rel" or any(k in d for k in ("rx1", "ry1", "rx2", "ry2")):
            return AreaAction(
                coord="rel",
                rx1=float(d.get("rx1", 0.0)), ry1=float(d.get("ry1", 0.0)),
                rx2=float(d.get("rx2", 1.0)), ry2=float(d.get("ry2", 1.0)),
                click=bool(d.get("click", False)),
                multiplier=max(1, int(d.get("multiplier", 1))),
                delay=delay,
                trigger=normalize_trigger(d.get("trigger", DEFAULT_TRIGGER)),
            )
        return AreaAction(
            coord="abs",
            x1=int(d.get("x1", 0)), y1=int(d.get("y1", 0)),
            x2=int(d.get("x2", 0)), y2=int(d.get("y2", 0)),
            click=bool(d.get("click", False)),
            multiplier=max(1, int(d.get("multiplier", 1))),
            delay=delay,
            trigger=normalize_trigger(d.get("trigger", DEFAULT_TRIGGER)),
        )


@dataclass
class KeyAction:
    type: ActionType = "key"
    kind: KeyKind = "keys"
    keys: List[str] = field(default_factory=list)  # e.g. ["Shift", "E"] or ["Escape"]
    mouse_button: Optional[MouseButtonName] = None  # if kind == mouse
    multiplier: int = 1
    delay: Delay = field(default_factory=lambda: Delay("fixed", 0.1, 0.1))
    press_mode: Literal["normal", "long"] = "normal"
    long_press: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "type": "key",
            "kind": self.kind,
            "keys": list(self.keys),
            "mouse_button": self.mouse_button,
            "multiplier": int(self.multiplier),
            "delay": self.delay.to_dict(),
            "press_mode": self.press_mode if self.press_mode in ("normal", "long") else "normal",
            "long_press": copy.deepcopy(self.long_press) if isinstance(self.long_press, dict) else {},
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "KeyAction":
        kind = d.get("kind", "keys")
        if kind not in ("keys", "mouse"):
            kind = "keys"
        mb = d.get("mouse_button", None)
        if mb not in (None, "left", "middle", "right"):
            mb = None
        if "delay" in d:
            delay = Delay.from_dict(d.get("delay", {}))
        else:
            delay = Delay("fixed", 0.1, 0.1)
        keys = d.get("keys", []) or []
        keys = [str(k) for k in keys][:10]
        press_mode = str(d.get("press_mode", "normal") or "normal").strip().lower()
        if press_mode not in ("normal", "long"):
            press_mode = "normal"
        long_press_raw = d.get("long_press", {})
        long_press = copy.deepcopy(long_press_raw) if isinstance(long_press_raw, dict) else {}
        return KeyAction(
            kind=kind,
            keys=keys,
            mouse_button=mb,
            multiplier=max(1, int(d.get("multiplier", 1))),
            delay=delay,
            press_mode=press_mode,
            long_press=long_press,
        )


@dataclass
class WaitAction:
    type: ActionType = "wait"
    delay: Delay = field(default_factory=lambda: Delay("fixed", 1.0, 1.0))

    def to_dict(self) -> Dict[str, Any]:
        return {"type": "wait", "delay": self.delay.to_dict()}

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "WaitAction":
        return WaitAction(delay=Delay.from_dict(d.get("delay", {})))


@dataclass
class WaitEventAction:
    type: ActionType = "wait_event"

    # legacy absolute
    x1: int = 0;
    y1: int = 0;
    x2: int = 0;
    y2: int = 0

    coord: Literal["abs", "rel"] = "rel"
    rx1: float = 0.0;
    ry1: float = 0.0;
    rx2: float = 1.0;
    ry2: float = 1.0

    expected_text: str = ""
    ocr_lang: str = "rus"
    poll: float = 1.0  # seconds between OCR checks

    def rect_abs(self) -> QRect:
        return QRect(QPoint(self.x1, self.y1), QPoint(self.x2, self.y2)).normalized()

    def rect_global(self, base: Optional[QRect]) -> QRect:
        if self.coord == "rel" and base and base.isValid() and base.width() >= 2 and base.height() >= 2:
            return rel_to_rect(base, self.rx1, self.ry1, self.rx2, self.ry2)
        return self.rect_abs()

    def resolve_target_rect_global(
        self,
        base: Optional[QRect] = None,
        debug: bool = False,
        debug_tag: str = "",
        dpr_override: Optional[float] = None,
    ) -> Optional[QRect]:
        probe = WordAreaAction(
            x1=int(self.x1),
            y1=int(self.y1),
            x2=int(self.x2),
            y2=int(self.y2),
            word=str(self.expected_text or ""),
            index=1,
            count=0,
            coord=self.coord,
            rx1=float(self.rx1),
            ry1=float(self.ry1),
            rx2=float(self.rx2),
            ry2=float(self.ry2),
            click=False,
            multiplier=1,
            delay=Delay("fixed", 0.1, 0.1),
            button="left",
            trigger=normalize_trigger(dict(DEFAULT_TRIGGER)),
            ocr_lang=str(self.ocr_lang or "rus"),
            search_infinite=True,
            search_max_tries=1,
            search_on_fail="retry",
        )
        return probe.resolve_target_rect_global(
            base=base,
            debug=debug,
            debug_tag=debug_tag,
            dpr_override=dpr_override,
        )

    @staticmethod
    def from_global(
        rect: QRect,
        base: QRect,
        *,
        expected_text: str = "",
        ocr_lang: str = "rus",
        poll: float = 1.0,
    ) -> "WaitEventAction":
        rx1, ry1, rx2, ry2 = rect_to_rel(base, rect)
        return WaitEventAction(
            coord="rel", rx1=rx1, ry1=ry1, rx2=rx2, ry2=ry2,
            expected_text=str(expected_text or ""),
            ocr_lang=str(ocr_lang or "rus"),
            poll=max(0.1, float(poll)),
        )

    def to_dict(self) -> Dict[str, Any]:
        out = {
            "type": "wait_event",
            "expected_text": str(self.expected_text),
            "ocr_lang": str(self.ocr_lang),
            "poll": float(self.poll),
        }
        if self.coord == "rel":
            out.update({"coord": "rel", "rx1": float(self.rx1), "ry1": float(self.ry1), "rx2": float(self.rx2),
                        "ry2": float(self.ry2)})
        else:
            out.update({"coord": "abs", "x1": int(self.x1), "y1": int(self.y1), "x2": int(self.x2), "y2": int(self.y2)})
        return out

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "WaitEventAction":
        coord = d.get("coord", None)
        poll_raw = d.get("poll", 1.0)
        try:
            poll = float(poll_raw)
        except Exception:
            poll = 1.0
        if poll < 0.1:
            poll = 0.1

        # Legacy migration: old "number" mode is converted to plain text matching.
        expected_text = str(d.get("expected_text", "") or "")
        if not expected_text.strip() and str(d.get("mode", "text") or "text") == "number":
            nv_raw = d.get("number_value", "")
            if isinstance(nv_raw, str):
                expected_text = nv_raw.strip()
            else:
                try:
                    expected_text = f"{float(nv_raw):g}"
                except Exception:
                    expected_text = ""

        base_kwargs = dict(
            expected_text=expected_text,
            ocr_lang=str(d.get("ocr_lang", "rus") or "rus"),
            poll=poll,
        )

        if coord == "rel" or any(k in d for k in ("rx1", "ry1", "rx2", "ry2")):
            return WaitEventAction(
                coord="rel",
                rx1=float(d.get("rx1", 0.0)), ry1=float(d.get("ry1", 0.0)),
                rx2=float(d.get("rx2", 1.0)), ry2=float(d.get("ry2", 1.0)),
                **base_kwargs,
            )
        return WaitEventAction(
            coord="abs",
            x1=int(d.get("x1", 0)), y1=int(d.get("y1", 0)),
            x2=int(d.get("x2", 0)), y2=int(d.get("y2", 0)),
            **base_kwargs,
        )


Action = Union[BaseAreaAction, AreaAction, KeyAction, WordAreaAction, WaitAction, WaitEventAction]


def _normalize_full_text(s: str) -> str:
    s = unicodedata.normalize("NFKC", (s or "")).strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("ё", "е")
    return s.casefold()


def ocr_text_in_rect(rect: QRect, lang: str = "rus", dpr_override: Optional[float] = None) -> Optional[str]:
    if not rect or not rect.isValid() or rect.width() < 2 or rect.height() < 2:
        return None
    available_langs = get_installed_ocr_languages(force_refresh=False)
    if not available_langs:
        return None
    lang_for_ocr = normalize_ocr_lang_spec(lang, available=available_langs)
    try:
        from PIL import Image, ImageGrab, ImageOps, ImageFilter, ImageStat
        import pytesseract
        from pytesseract import TesseractNotFoundError
    except Exception:
        return None

    dpr = 1.0
    if dpr_override is not None:
        try:
            dpr = float(dpr_override)
            if dpr <= 0:
                dpr = 1.0
        except Exception:
            dpr = 1.0
    else:
        try:
            scr = QGuiApplication.screenAt(rect.center())
            if scr:
                dpr = float(scr.devicePixelRatio())
        except Exception:
            dpr = 1.0

    bbox = (
        int(round(rect.left() * dpr)),
        int(round(rect.top() * dpr)),
        int(round((rect.right() + 1) * dpr)),
        int(round((rect.bottom() + 1) * dpr)),
    )

    try:
        try:
            img = ImageGrab.grab(bbox=bbox, all_screens=True)
        except TypeError:
            img = ImageGrab.grab(bbox=bbox)
    except Exception:
        return None

    def otsu_threshold(gray_img) -> int:
        hist = gray_img.histogram()
        total = sum(hist)
        if total <= 0:
            return 128
        sum_all = 0
        for i, h in enumerate(hist):
            sum_all += i * h
        sum_b = 0
        w_b = 0
        var_max = -1.0
        thr = 128
        for t in range(256):
            w_b += hist[t]
            if w_b == 0:
                continue
            w_f = total - w_b
            if w_f == 0:
                break
            sum_b += t * hist[t]
            m_b = sum_b / w_b
            m_f = (sum_all - sum_b) / w_f
            var_between = w_b * w_f * (m_b - m_f) * (m_b - m_f)
            if var_between > var_max:
                var_max = var_between
                thr = t
        return int(thr)

    try:
        orig = img.convert("RGB")
        ow, oh = orig.size
        min_dim = min(ow, oh)
        if min_dim < 220:
            scale = 4
        elif min_dim < 420:
            scale = 3
        elif min_dim < 900:
            scale = 2
        else:
            scale = 1
        resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)
        up = orig.resize((max(1, ow * scale), max(1, oh * scale)), resampling)
        up = ImageOps.grayscale(up)
        up = ImageOps.autocontrast(up)
        try:
            mean = ImageStat.Stat(up).mean[0]
            if mean < 110:
                up = ImageOps.invert(up)
        except Exception:
            pass
        up = up.filter(ImageFilter.UnsharpMask(radius=1.2, percent=180, threshold=2))
        thr = otsu_threshold(up)
        up = up.point(lambda p: 255 if p > thr else 0)
    except Exception:
        return None

    try:
        txt = pytesseract.image_to_string(up, lang=lang_for_ocr, config="--oem 3 --psm 6")
    except TesseractNotFoundError:
        return None
    except Exception:
        return None

    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


@dataclass
class Record:
    name: str = "Новая запись"
    actions: List[Dict[str, Any]] = field(default_factory=list)
    move_mouse: bool = True
    repeat: RepeatSettings = field(default_factory=RepeatSettings)
    bind_to_process: bool = False
    bound_exe: str = ""
    bound_exe_override: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "actions": self.actions,
            "move_mouse": bool(self.move_mouse),
            "repeat": self.repeat.to_dict(),
            "bind_to_process": bool(self.bind_to_process),
            "bound_exe": str(self.bound_exe or "").strip(),
            "bound_exe_override": str(self.bound_exe_override or "").strip(),
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Record":
        name = str(d.get("name", "Запись"))
        actions = d.get("actions", []) or []
        if not isinstance(actions, list):
            actions = []
        move_mouse_val = d.get("move_mouse", True)
        move_mouse = True if move_mouse_val is None else bool(move_mouse_val)
        repeat = RepeatSettings.from_dict(d.get("repeat", {}))
        bind_to_process = bool(d.get("bind_to_process", False))
        bound_exe = str(d.get("bound_exe", "") or "").strip()
        bound_exe_override = str(d.get("bound_exe_override", "") or "").strip()
        if bind_to_process and (not bound_exe) and bound_exe_override:
            bound_exe = bound_exe_override
            bound_exe_override = ""
        return Record(
            name=name,
            actions=actions,
            move_mouse=move_mouse,
            repeat=repeat,
            bind_to_process=bind_to_process,
            bound_exe=bound_exe,
            bound_exe_override=bound_exe_override,
        )


def action_from_dict(d: Dict[str, Any]) -> Action:
    t = d.get("type")
    if t == "area":
        return AreaAction.from_dict(d)
    if t == "area_word":
        return WordAreaAction.from_dict(d)
    if t == "wait":
        return WaitAction.from_dict(d)
    if t == "base_area":
        return BaseAreaAction.from_dict(d)
    if t == "wait_event":
        return WaitEventAction.from_dict(d)

    return KeyAction.from_dict(d)


def action_to_display(a: Action) -> Tuple[str, str]:
    if isinstance(a, BaseAreaAction):
        r = a.rect()
        return (
            "Опорная область",
            f"рамка: ({r.left()}, {r.top()}) → ({r.right()}, {r.bottom()}); размер: {r.width()}×{r.height()}",
        )
    if isinstance(a, WaitAction):
        if a.delay.mode == "range":
            return ("Ожидание", f"длительность: {a.delay.a:.3f}–{a.delay.b:.3f}с (случайно)")
        return ("Ожидание", f"длительность: {a.delay.a:.3f}с")
    if isinstance(a, WaitEventAction):
        cond = f"текст \"{a.expected_text}\""
        if a.coord == "rel":
            return ("Ожидание", f"{cond}; зона: rel=({a.rx1:.3f},{a.ry1:.3f})→({a.rx2:.3f},{a.ry2:.3f})")
        r = a.rect_abs()
        return ("Ожидание", f"{cond}; зона: ({r.left()},{r.top()})→({r.right()},{r.bottom()})")
    if isinstance(a, WordAreaAction):
        count_val = int(getattr(a, "count", 0))
        count_txt = f"{a.index}/{count_val}" if count_val > 0 else str(a.index)
        if a.coord == "rel":
            return (
                "Область",
                f"текст \"{a.word}\"; вхождение: {count_txt}; зона: rel=({a.rx1:.3f},{a.ry1:.3f})→({a.rx2:.3f},{a.ry2:.3f})",
            )
        r = a.search_rect_abs()
        return (
            "Область",
            f"текст \"{a.word}\"; вхождение: {count_txt}; зона: ({r.left()},{r.top()})→({r.right()},{r.bottom()})",
        )

    if isinstance(a, AreaAction):
        if a.coord == "rel":
            return ("Область", f"зона: rel=({a.rx1:.3f},{a.ry1:.3f})→({a.rx2:.3f},{a.ry2:.3f})")
        r = a.rect_abs()
        return ("Область", f"зона: ({r.left()},{r.top()})→({r.right()},{r.bottom()}); размер: {r.width()}×{r.height()}")

    else:
        if a.kind == "mouse":
            btn = {"left": "ЛКМ", "middle": "СКМ", "right": "ПКМ"}.get(a.mouse_button or "", "Мышь")
            base = f"мышь: {btn}"
        else:
            base = f"клавиши: {' + '.join(a.keys) if a.keys else '(пусто)'}"
        return ("Нажатие", base)

