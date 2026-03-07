from __future__ import annotations

from typing import Any, Dict


class UnoObject:
    """Holds a UNO object and caches queried interfaces."""

    def __init__(self, obj: Any) -> None:
        self._obj = obj
        self._iface_cache: Dict[str, Any] = {}

    def __getattr__(self, name: str) -> Any:  # pragma: no cover - passthrough
        # Wrapperで未定義の属性は生のUNOオブジェクトへ委譲する。
        return getattr(self._obj, name)

    def __setattr__(self, name: str, value: Any) -> None:  # pragma: no cover - passthrough
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        # クラス側に setter 付きの property があれば優先する
        cls_attr = getattr(type(self), name, None)
        if isinstance(cls_attr, property) and cls_attr.fset is not None:
            cls_attr.fset(self, value)  # type: ignore[misc]
            return
        # 通常の属性として持っていればそのまま設定
        if hasattr(type(self), name) or name in self.__dict__:
            object.__setattr__(self, name, value)
            return
        # それ以外は UNO オブジェクトへ委譲
        try:
            setattr(self._obj, name, value)
        except AttributeError as exc:
            raise AttributeError(name) from exc

    def iface(self, name: str) -> Any:
        """Query and memoize a UNO interface by name."""
        if name not in self._iface_cache:
            query = getattr(self._obj, "queryInterface", None)
            if query is None:
                raise AttributeError("UNO object missing queryInterface")
            iface_obj: Any = None

            # Prefer UNO type resolution when available (real UNO objects),
            # but fall back to the provided name for lightweight test doubles.
            try:
                import uno  # type: ignore

                try:
                    iface_type = uno.getTypeByName(name)
                    iface_obj = query(iface_type)
                except Exception:
                    iface_obj = None
            except Exception:
                iface_obj = None

            if iface_obj is None:
                iface_obj = query(name)

            self._iface_cache[name] = iface_obj
        return self._iface_cache[name]

    @property
    def raw(self) -> Any:
        """Expose the wrapped UNO object when direct access is needed."""
        return self._obj

    def queryInterface(self, iface: Any) -> Any:
        """Delegate queryInterface to the wrapped UNO object when present."""
        target = getattr(self._obj, "queryInterface", None)
        if target is None:
            raise AttributeError("UNO object missing queryInterface")
        return target(iface)

