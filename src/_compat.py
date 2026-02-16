"""Compatibility helpers for legacy src.* imports."""

from importlib import import_module
from types import ModuleType
from typing import Any, Dict


def re_export(target_module: str, namespace: Dict[str, Any]) -> ModuleType:
    module = import_module(target_module)

    exported = {
        name: getattr(module, name)
        for name in dir(module)
        if not (name.startswith("__") and name.endswith("__"))
    }
    namespace.update(exported)

    if "__all__" in exported and isinstance(exported["__all__"], list):
        namespace["__all__"] = list(exported["__all__"])
    else:
        namespace["__all__"] = [name for name in exported.keys() if name != "__all__"]

    namespace["__doc__"] = getattr(module, "__doc__", namespace.get("__doc__"))
    namespace["_TARGET_MODULE"] = module
    return module
