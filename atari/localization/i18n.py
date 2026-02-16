"""Runtime UI localization helpers for Russian/English."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

DEFAULT_LANGUAGE = "en"
SUPPORTED_LANGUAGES: Tuple[str, ...] = ("en", "ru")

_PLACEHOLDER_RE = re.compile(r"\{(\d+)\}")
_CYR_RE = re.compile(r"[\u0400-\u04FF]")
_CURRENT_LANGUAGE = DEFAULT_LANGUAGE
_DATA_LOADED = False
_CACHE: Dict[Tuple[str, str], str] = {}

_RU_TO_EN_EXACT: Dict[str, str] = {}
_EN_TO_RU_EXACT: Dict[str, str] = {}
_RU_TO_EN_FRAG: List[Tuple[str, str]] = []
_EN_TO_RU_FRAG: List[Tuple[str, str]] = []

_MANUAL_RU_TO_EN: Dict[str, str] = {
    "\u042f\u0437\u044b\u043a \u043f\u0440\u0438\u043b\u043e\u0436\u0435\u043d\u0438\u044f": "App language",
    "\u041f\u0440\u043e\u0438\u0433\u0440\u0430\u0442\u044c": "Play",
    "\u25b6 \u041f\u0440\u043e\u0438\u0433\u0440\u0430\u0442\u044c (F4)": "\u25b6 Play (F4)",
    "\u041e\u0436\u0438\u0434\u0430\u043d\u0438\u0435": "Wait",
    "\u041e\u0431\u043b\u0430\u0441\u0442\u044c": "Area",
    "\u041d\u0430\u0436\u0430\u0442\u0438\u0435": "Press",
    "\u0417\u0430\u043f\u0438\u0441\u0438": "Records",
    "\u0418\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0435": "Edit",
    "\u0423\u0434\u0430\u043b\u0435\u043d\u0438\u0435": "Delete",
    "\u041e\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c\u0441\u044f": "Stop",
    "\u0423\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c \u0432\u043d\u0438\u043c\u0430\u043d\u0438\u0435": "Set focus",
    "\u0412\u044b\u0432\u0435\u0441\u0442\u0438 \u043e\u0448\u0438\u0431\u043a\u0443": "Show error",
    "\u041e\u0447\u0438\u0441\u0442\u0438\u0442\u044c \u043d\u0435 \u0438\u0437\u0431\u0440\u0430\u043d\u043d\u044b\u0435": "Clear non-favorites",
    "\u041a\u043e\u043b-\u0432\u043e": "Count",
    "\u0417\u0430\u043f\u0443\u0441\u043a": "Start",
}


@dataclass(frozen=True)
class _TemplateRule:
    pattern: re.Pattern[str]
    target_template: str


_RU_TO_EN_TPL: List[_TemplateRule] = []
_EN_TO_RU_TPL: List[_TemplateRule] = []

_QT_HOOKS_INSTALLED = False
_ORIGINAL_QT: Dict[Tuple[str, str], Any] = {}


def normalize_language(lang: Optional[str]) -> str:
    code = str(lang or "").strip().lower()
    if code in SUPPORTED_LANGUAGES:
        return code
    return DEFAULT_LANGUAGE


def get_language() -> str:
    return _CURRENT_LANGUAGE


def set_language(lang: Optional[str]) -> str:
    global _CURRENT_LANGUAGE
    _CURRENT_LANGUAGE = normalize_language(lang)
    _CACHE.clear()
    return _CURRENT_LANGUAGE


def language_choices() -> List[Tuple[str, str]]:
    # Keep native language names.
    return [("en", "English"), ("ru", "Русский")]


def _data_path() -> Path:
    return Path(__file__).with_name("i18n_data.json")


def _render_template(template: str, values: Dict[int, str]) -> str:
    def _replace(m: re.Match[str]) -> str:
        idx = int(m.group(1))
        return values.get(idx, m.group(0))

    return _PLACEHOLDER_RE.sub(_replace, template)


def _compile_template(source_template: str, target_template: str) -> _TemplateRule:
    parts: List[str] = []
    seen: set[int] = set()
    pos = 0
    for m in _PLACEHOLDER_RE.finditer(source_template):
        parts.append(re.escape(source_template[pos : m.start()]))
        idx = int(m.group(1))
        name = f"v{idx}"
        if idx in seen:
            parts.append(f"(?P={name})")
        else:
            parts.append(f"(?P<{name}>.+?)")
            seen.add(idx)
        pos = m.end()
    parts.append(re.escape(source_template[pos:]))
    regex = re.compile("^" + "".join(parts) + "$", re.DOTALL)
    return _TemplateRule(pattern=regex, target_template=target_template)


def _template_weight(item: Tuple[str, str]) -> Tuple[int, int]:
    src = item[0]
    # Prefer more specific patterns first.
    return (len(src), -src.count("{"))


def _fragment_weight(item: Tuple[str, str]) -> int:
    return len(item[0])


def _preserve_edge_spaces(src: str, dst: str) -> str:
    if not src or not dst:
        return dst
    l_src = len(src) - len(src.lstrip(" "))
    r_src = len(src) - len(src.rstrip(" "))
    core = dst.strip(" ")
    return (" " * l_src) + core + (" " * r_src)


def _reverse_source_score(src: str) -> Tuple[int, int, int, int]:
    # Prefer readable native text over corrupted placeholders.
    has_cyr = 1 if _CYR_RE.search(src) else 0
    qmarks = src.count("?")
    return (has_cyr, -qmarks, 1 if qmarks == 0 else 0, len(src))


def _load_data() -> None:
    global _DATA_LOADED
    if _DATA_LOADED:
        return

    try:
        payload = json.loads(_data_path().read_text(encoding="utf-8"))
    except Exception:
        payload = {}

    ru_to_en_exact = payload.get("ru_to_en_exact", {}) or {}
    ru_to_en_tpl = payload.get("ru_to_en_templates", {}) or {}

    if not isinstance(ru_to_en_exact, dict):
        ru_to_en_exact = {}
    if not isinstance(ru_to_en_tpl, dict):
        ru_to_en_tpl = {}

    _RU_TO_EN_EXACT.clear()
    _EN_TO_RU_EXACT.clear()
    _RU_TO_EN_FRAG.clear()
    _EN_TO_RU_FRAG.clear()
    _RU_TO_EN_TPL.clear()
    _EN_TO_RU_TPL.clear()

    for k, v in ru_to_en_exact.items():
        src = str(k)
        dst = _preserve_edge_spaces(src, str(v))
        _RU_TO_EN_EXACT[src] = dst

    for src, dst in _MANUAL_RU_TO_EN.items():
        _RU_TO_EN_EXACT[src] = dst

    _EN_TO_RU_EXACT.clear()
    reverse_best: Dict[str, Tuple[str, Tuple[int, int, int, int]]] = {}
    for src, dst in _RU_TO_EN_EXACT.items():
        cand_score = _reverse_source_score(src)
        prev = reverse_best.get(dst)
        if prev is None or cand_score > prev[1]:
            reverse_best[dst] = (src, cand_score)
    for dst, pair in reverse_best.items():
        _EN_TO_RU_EXACT[dst] = pair[0]

    tpl_items = [(str(k), _preserve_edge_spaces(str(k), str(v))) for k, v in ru_to_en_tpl.items()]
    tpl_items.sort(key=_template_weight, reverse=True)
    for src_tpl, dst_tpl in tpl_items:
        _RU_TO_EN_TPL.append(_compile_template(src_tpl, dst_tpl))
        _EN_TO_RU_TPL.append(_compile_template(dst_tpl, src_tpl))

    def _should_use_fragment(src: str, dst: str) -> bool:
        if not src or not dst or src == dst:
            return False
        if "\n" in src or "\n" in dst:
            return False
        if "{" in src or "}" in src:
            return False
        if len(src) < 2 or len(src) > 72:
            return False
        return True

    for src, dst in _RU_TO_EN_EXACT.items():
        if _should_use_fragment(src, dst):
            _RU_TO_EN_FRAG.append((src, dst))
            _EN_TO_RU_FRAG.append((dst, src))

    _RU_TO_EN_FRAG.sort(key=_fragment_weight, reverse=True)
    _EN_TO_RU_FRAG.sort(key=_fragment_weight, reverse=True)

    _DATA_LOADED = True


def _translate_by_templates(
    text: str,
    rules: Iterable[_TemplateRule],
    target_language: Optional[str] = None,
) -> Optional[str]:
    for rule in rules:
        m = rule.pattern.match(text)
        if not m:
            continue
        values: Dict[int, str] = {}
        for key, val in m.groupdict().items():
            if val is None:
                continue
            if not key.startswith("v"):
                continue
            try:
                idx = int(key[1:])
            except Exception:
                continue
            if target_language and val != text:
                values[idx] = str(translate_text(val, target_language))
            else:
                values[idx] = val
        return _render_template(rule.target_template, values)
    return None


def _replace_fragments(text: str, replacements: Iterable[Tuple[str, str]]) -> str:
    out = text
    for src, dst in replacements:
        if src in out:
            out = out.replace(src, dst)
    return out


def translate_text(value: Any, target_language: Optional[str] = None) -> Any:
    if not isinstance(value, str) or not value:
        return value

    _load_data()
    lang = normalize_language(target_language or _CURRENT_LANGUAGE)
    key = (lang, value)
    cached = _CACHE.get(key)
    if cached is not None:
        return cached

    if lang == "ru":
        out = _EN_TO_RU_EXACT.get(value)
        if out is None:
            out = _translate_by_templates(value, _EN_TO_RU_TPL, target_language="ru")
        if out is None:
            frag = _replace_fragments(value, _EN_TO_RU_FRAG)
            out = frag if frag != value else None
        if out is None:
            out = value
    else:
        out = _RU_TO_EN_EXACT.get(value)
        if out is None:
            out = _translate_by_templates(value, _RU_TO_EN_TPL, target_language="en")
        if out is None:
            frag = _replace_fragments(value, _RU_TO_EN_FRAG)
            out = frag if frag != value else None
        if out is None:
            out = value

    _CACHE[key] = out
    return out


def tr(text: str) -> str:
    return str(translate_text(text))


def _patch_qt_method(owner: Any, name: str, wrapper_builder) -> None:
    key = (getattr(owner, "__name__", str(owner)), name)
    if key in _ORIGINAL_QT:
        return
    original = getattr(owner, name, None)
    if original is None:
        return
    _ORIGINAL_QT[key] = original
    setattr(owner, name, wrapper_builder(original))


def install_qt_translation_hooks() -> None:
    global _QT_HOOKS_INSTALLED
    if _QT_HOOKS_INSTALLED:
        return

    try:
        from PySide6.QtGui import QAction
        from PySide6.QtWidgets import (
            QAbstractButton,
            QCheckBox,
            QComboBox,
            QFileDialog,
            QGroupBox,
            QInputDialog,
            QLabel,
            QLineEdit,
            QListWidgetItem,
            QMessageBox,
            QMenu,
            QPushButton,
            QTableWidget,
            QTableWidgetItem,
            QToolButton,
            QWidget,
        )
    except Exception:
        return

    def _wrap_text_arg0(original):
        def _patched(self, text, *args, **kwargs):
            return original(self, translate_text(text), *args, **kwargs)

        return _patched

    def _wrap_list_arg0(original):
        def _patched(self, labels, *args, **kwargs):
            if isinstance(labels, (list, tuple)):
                labels = [translate_text(x) for x in labels]
            return original(self, labels, *args, **kwargs)

        return _patched

    def _wrap_qtable_item_init(original):
        def _patched(self, *args, **kwargs):
            a = list(args)
            if a and isinstance(a[0], str):
                a[0] = translate_text(a[0])
            if len(a) >= 2 and not isinstance(a[0], str) and isinstance(a[1], str):
                a[1] = translate_text(a[1])
            if isinstance(kwargs.get("text"), str):
                kwargs["text"] = translate_text(kwargs["text"])
            return original(self, *tuple(a), **kwargs)

        return _patched

    def _wrap_translate_text_init(original):
        def _patched(self, *args, **kwargs):
            a = list(args)
            for idx in (0, 1):
                if idx < len(a) and isinstance(a[idx], str):
                    a[idx] = translate_text(a[idx])
            return original(self, *tuple(a), **kwargs)

        return _patched

    def _wrap_qlist_item_init(original):
        def _patched(self, *args, **kwargs):
            a = list(args)
            if a and isinstance(a[0], str):
                a[0] = translate_text(a[0])
            if len(a) >= 2 and not isinstance(a[0], str) and isinstance(a[1], str):
                a[1] = translate_text(a[1])
            return original(self, *tuple(a), **kwargs)

        return _patched

    def _wrap_menu_add_action(original):
        def _patched(self, *args, **kwargs):
            a = list(args)
            if a and isinstance(a[0], str):
                a[0] = translate_text(a[0])
            if len(a) >= 2 and not isinstance(a[0], str) and isinstance(a[1], str):
                a[1] = translate_text(a[1])
            if isinstance(kwargs.get("text"), str):
                kwargs["text"] = translate_text(kwargs["text"])
            return original(self, *tuple(a), **kwargs)

        return _patched

    def _wrap_combo_add_item(original):
        def _patched(self, *args, **kwargs):
            a = list(args)
            if a and isinstance(a[0], str):
                a[0] = translate_text(a[0])
            if len(a) >= 2 and not isinstance(a[0], str) and isinstance(a[1], str):
                a[1] = translate_text(a[1])
            return original(self, *tuple(a), **kwargs)

        return _patched

    def _wrap_combo_add_items(original):
        def _patched(self, items, *args, **kwargs):
            if isinstance(items, (list, tuple)):
                items = [translate_text(x) for x in items]
            return original(self, items, *args, **kwargs)

        return _patched

    def _wrap_qmessagebox_static(original):
        def _patched(parent, title, text, *args, **kwargs):
            return original(parent, translate_text(title), translate_text(text), *args, **kwargs)

        return staticmethod(_patched)

    def _wrap_qfiledialog_static(original):
        def _patched(*args, **kwargs):
            a = list(args)
            if len(a) >= 2 and isinstance(a[1], str):
                a[1] = translate_text(a[1])
            if len(a) >= 4 and isinstance(a[3], str):
                a[3] = translate_text(a[3])
            if isinstance(kwargs.get("caption"), str):
                kwargs["caption"] = translate_text(kwargs["caption"])
            if isinstance(kwargs.get("filter"), str):
                kwargs["filter"] = translate_text(kwargs["filter"])
            return original(*tuple(a), **kwargs)

        return staticmethod(_patched)

    _patch_qt_method(QLabel, "setText", _wrap_text_arg0)
    _patch_qt_method(QAbstractButton, "setText", _wrap_text_arg0)
    _patch_qt_method(QGroupBox, "setTitle", _wrap_text_arg0)
    _patch_qt_method(QWidget, "setWindowTitle", _wrap_text_arg0)
    _patch_qt_method(QAction, "setText", _wrap_text_arg0)
    _patch_qt_method(QLineEdit, "setPlaceholderText", _wrap_text_arg0)
    _patch_qt_method(QInputDialog, "setLabelText", _wrap_text_arg0)
    _patch_qt_method(QTableWidget, "setHorizontalHeaderLabels", _wrap_list_arg0)
    _patch_qt_method(QTableWidgetItem, "__init__", _wrap_qtable_item_init)
    _patch_qt_method(QTableWidgetItem, "setText", _wrap_text_arg0)
    _patch_qt_method(QListWidgetItem, "__init__", _wrap_qlist_item_init)
    _patch_qt_method(QListWidgetItem, "setText", _wrap_text_arg0)
    _patch_qt_method(QMenu, "addAction", _wrap_menu_add_action)
    _patch_qt_method(QComboBox, "addItem", _wrap_combo_add_item)
    _patch_qt_method(QComboBox, "addItems", _wrap_combo_add_items)
    _patch_qt_method(QComboBox, "setItemText", _wrap_text_arg0)
    _patch_qt_method(QLabel, "__init__", _wrap_translate_text_init)
    _patch_qt_method(QPushButton, "__init__", _wrap_translate_text_init)
    _patch_qt_method(QToolButton, "__init__", _wrap_translate_text_init)
    _patch_qt_method(QCheckBox, "__init__", _wrap_translate_text_init)
    _patch_qt_method(QGroupBox, "__init__", _wrap_translate_text_init)
    _patch_qt_method(QAction, "__init__", _wrap_translate_text_init)

    _patch_qt_method(QMessageBox, "information", _wrap_qmessagebox_static)
    _patch_qt_method(QMessageBox, "warning", _wrap_qmessagebox_static)
    _patch_qt_method(QMessageBox, "critical", _wrap_qmessagebox_static)
    _patch_qt_method(QMessageBox, "question", _wrap_qmessagebox_static)
    _patch_qt_method(QFileDialog, "getOpenFileName", _wrap_qfiledialog_static)
    _patch_qt_method(QFileDialog, "getOpenFileNames", _wrap_qfiledialog_static)
    _patch_qt_method(QFileDialog, "getSaveFileName", _wrap_qfiledialog_static)
    _patch_qt_method(QFileDialog, "getExistingDirectory", _wrap_qfiledialog_static)

    _QT_HOOKS_INSTALLED = True


def _retranslate_action_tree(actions: Iterable[Any]) -> None:
    for act in actions:
        try:
            act.setText(act.text())
        except Exception:
            pass
        try:
            submenu = act.menu()
        except Exception:
            submenu = None
        if submenu is not None:
            retranslate_widget_tree(submenu)


def retranslate_widget_tree(root: Any) -> None:
    """Re-apply text values for all known widgets to update active language."""
    if root is None:
        return

    try:
        from PySide6.QtWidgets import (
            QComboBox,
            QGroupBox,
            QLineEdit,
            QListWidget,
            QMenu,
            QTableWidget,
            QWidget,
        )
    except Exception:
        return

    queue: List[Any] = []
    try:
        queue.append(root)
        if isinstance(root, QWidget):
            queue.extend(root.findChildren(QWidget))
    except Exception:
        pass

    for widget in queue:
        # Window title
        try:
            title = widget.windowTitle()
            if isinstance(title, str):
                widget.setWindowTitle(title)
        except Exception:
            pass

        # Main text
        try:
            text = widget.text()
            if isinstance(text, str):
                widget.setText(text)
        except Exception:
            pass

        if isinstance(widget, QGroupBox):
            try:
                widget.setTitle(widget.title())
            except Exception:
                pass

        if isinstance(widget, QLineEdit):
            try:
                widget.setPlaceholderText(widget.placeholderText())
            except Exception:
                pass

        if isinstance(widget, QComboBox):
            try:
                for i in range(widget.count()):
                    widget.setItemText(i, widget.itemText(i))
            except Exception:
                pass

        if isinstance(widget, QListWidget):
            try:
                for i in range(widget.count()):
                    it = widget.item(i)
                    if it is not None:
                        it.setText(it.text())
            except Exception:
                pass

        if isinstance(widget, QTableWidget):
            try:
                labels: List[str] = []
                for c in range(widget.columnCount()):
                    item = widget.horizontalHeaderItem(c)
                    labels.append(item.text() if item is not None else "")
                if labels:
                    widget.setHorizontalHeaderLabels(labels)
            except Exception:
                pass
            try:
                for r in range(widget.rowCount()):
                    for c in range(widget.columnCount()):
                        item = widget.item(r, c)
                        if item is not None:
                            item.setText(item.text())
            except Exception:
                pass

        if isinstance(widget, QMenu):
            try:
                _retranslate_action_tree(widget.actions())
            except Exception:
                pass

        # ToolButton/QPushButton menus are not always children in hierarchy.
        try:
            menu = widget.menu()
        except Exception:
            menu = None
        if menu is not None:
            retranslate_widget_tree(menu)

    # Root can be QMenu (not QWidget child traversal in some cases).
    try:
        if hasattr(root, "actions"):
            _retranslate_action_tree(root.actions())
    except Exception:
        pass
