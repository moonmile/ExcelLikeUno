from __future__ import annotations

from typing import Any, Dict


class UnoObject:
    """Holds a UNO object and caches queried interfaces."""

    def __init__(self, obj: Any) -> None:
        self._obj = obj
        self._iface_cache: Dict[str, Any] = {}

    def iface(self, name: str) -> Any:
        """Query and memoize a UNO interface by name."""
        if name not in self._iface_cache:
            query = getattr(self._obj, "queryInterface", None)
            if query is None:
                raise AttributeError("UNO object missing queryInterface")
            self._iface_cache[name] = query(name)
        return self._iface_cache[name]

    @property
    def raw(self) -> Any:
        """Expose the wrapped UNO object when direct access is needed."""
        return self._obj
