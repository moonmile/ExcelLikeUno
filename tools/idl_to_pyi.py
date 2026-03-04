from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

# Simple IDL -> PYI generator specialized for LibreOffice offapi tree.
# Goal: produce Protocol-based stubs with inheritance info and mirror the IDL package layout.

PRIMITIVE_MAP = {
    "void": "None",
    "any": "Any",
    "byte": "int",
    "short": "int",
    "long": "int",
    "hyper": "int",
    "unsigned short": "int",
    "unsigned long": "int",
    "float": "float",
    "double": "float",
    "boolean": "bool",
    "char": "str",
    "string": "str",
}


@dataclass
class Method:
    name: str
    return_type: str
    params: list[tuple[str, str]] = field(default_factory=list)


@dataclass
class Attribute:
    name: str
    type_name: str


@dataclass
class Definition:
    kind: str  # interface | struct | enum | service
    name: str
    parents: list[str] = field(default_factory=list)
    methods: list[Method] = field(default_factory=list)
    attributes: list[Attribute] = field(default_factory=list)
    enum_members: list[str] = field(default_factory=list)
    struct_fields: list[Attribute] = field(default_factory=list)


def strip_comments(text: str) -> str:
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)
    text = re.sub(r"//.*", "", text)
    return text


def map_type(raw: str) -> str:
    raw = raw.strip()
    raw = re.sub(r"\s+", " ", raw)
    raw = raw.replace("[", "").replace("]", "")
    if raw.startswith("sequence<") and raw.endswith(">"):
        inner = raw[len("sequence<") : -1].strip()
        return f"list[{map_type(inner)}]"
    # remove leading qualifiers like [in], const, etc.
    raw = re.sub(r"\b(in|out|inout|const)\b", "", raw).strip()
    # collapse spaces
    raw = re.sub(r"\s+", " ", raw)
    if raw in PRIMITIVE_MAP:
        return PRIMITIVE_MAP[raw]
    # fully qualified :: names -> dotted import will target this name, keep last component
    if "::" in raw:
        return raw.split("::")[-1]
    return raw or "Any"


def parse_methods(body: str) -> list[Method]:
    methods: list[Method] = []
    pattern = re.compile(r"([A-Za-z_][\w:<>\s]*)\s+([A-Za-z_][\w]*)\s*\(([^)]*)\)\s*;")
    for match in pattern.finditer(body):
        ret_raw, name, params_raw = match.groups()
        params: list[tuple[str, str]] = []
        params_raw = params_raw.strip()
        if params_raw:
            for param in params_raw.split(','):
                param = param.strip()
                if not param:
                    continue
                parts = param.rsplit(' ', 1)
                if len(parts) == 2:
                    p_type_raw, p_name = parts
                else:
                    p_type_raw, p_name = param, f"param{len(params)}"
                params.append((map_type(p_type_raw), p_name))
        methods.append(Method(name=name, return_type=map_type(ret_raw), params=params))
    return methods


def parse_attributes(body: str) -> list[Attribute]:
    attrs: list[Attribute] = []
    # 1) attribute/property keyword forms (interfaces often use this)
    pattern_attr = re.compile(
        r"(?:\[[^\]]*\]\s*)?(?:readonly\s+)?(?:attribute|property)\s+([A-Za-z_][\w:<>,\s]*)\s+([A-Za-z_][\w]*)\s*;"
    )
    for match in pattern_attr.finditer(body):
        t_raw, name = match.groups()
        attrs.append(Attribute(name=name, type_name=map_type(t_raw)))

    # 2) UNO service [property] annotations without an explicit keyword
    pattern_prop_bracket = re.compile(
        r"\[[^\]]*property[^\]]*\]\s*([A-Za-z_][\w:<>,\s]*)\s+([A-Za-z_][\w]*)\s*;"
    )
    for match in pattern_prop_bracket.finditer(body):
        t_raw, name = match.groups()
        attrs.append(Attribute(name=name, type_name=map_type(t_raw)))
    return attrs


def parse_struct_fields(body: str) -> list[Attribute]:
    fields: list[Attribute] = []
    pattern = re.compile(r"([A-Za-z_][\w:<>\s]*)\s+([A-Za-z_][\w]*)\s*;")
    for match in pattern.finditer(body):
        t_raw, name = match.groups()
        fields.append(Attribute(name=name, type_name=map_type(t_raw)))
    return fields


def parse_enum_members(body: str) -> list[str]:
    cleaned = body.replace("\n", " ")
    cleaned = cleaned.replace("{", " ").replace("}", " ")
    parts = [p.strip() for p in cleaned.split(',')]
    return [p for p in parts if p and p != ';']


def parse_definitions(text: str) -> list[Definition]:
    defs: list[Definition] = []
    text = strip_comments(text)
    # Simplify whitespace
    text = re.sub(r"\s+", " ", text)

    def parse_block(kind: str, match: re.Match) -> Definition:
        name = match.group(1)
        parents_raw = ""
        body = ""
        if kind == "enum":
            body = match.group(2) or ""
        else:
            parents_raw = match.group(2) or ""
            body = match.group(3) or ""
        parents = [p.strip() for p in parents_raw.split(',') if p.strip()]
        # Services can declare additional parents inside the body via `service ...;` or `interface ...;`
        if kind == "service" and body:
            extra = re.findall(r"(?:service|interface)\s+([A-Za-z_][\w:]*)\s*;", body)
            for parent in extra:
                if parent and parent not in parents:
                    parents.append(parent)
        definition = Definition(kind=kind, name=name, parents=parents)
        if kind in {"interface", "service"}:
            definition.methods = parse_methods(body)
            definition.attributes = parse_attributes(body)
        elif kind == "struct":
            definition.struct_fields = parse_struct_fields(body)
        elif kind == "enum":
            definition.enum_members = parse_enum_members(body)
        return definition

    patterns = {
        # parents clause may contain multiple comma-separated names; match until the next '{' non-greedily.
        "interface": re.compile(r"interface\s+([A-Za-z_][\w]*)\s*(?::\s*([^{}]+?))?\s*{(.*?)}", re.S),
        "service": re.compile(r"service\s+([A-Za-z_][\w]*)\s*(?:extends\s*([^{}]+?))?\s*{(.*?)}", re.S),
        "struct": re.compile(r"struct\s+([A-Za-z_][\w]*)\s*(?::\s*([^{}]+?))?\s*{(.*?)}", re.S),
        "enum": re.compile(r"enum\s+([A-Za-z_][\w]*)\s*{(.*?)}", re.S),
    }

    for kind, pat in patterns.items():
        for match in pat.finditer(text):
            defs.append(parse_block(kind, match))
    return defs


def parent_imports(parents: list[str]) -> list[str]:
    imports: list[str] = []
    for parent in parents:
        if "::" in parent:
            parts = parent.split("::")
            module_path = ".".join(parts[:-1])
            name = parts[-1]
            imports.append(f"from {module_path} import {name}")
    return imports


def render_definition(defn: Definition) -> str:
    lines: list[str] = []
    if defn.kind in {"interface", "service"}:
        base_raw = [p for p in defn.parents if p]
        base_names: list[str] = []
        for b in base_raw:
            name = b.split("::")[-1]
            if name not in base_names:
                base_names.append(name)
        if "Protocol" not in base_names:
            base_names.append("Protocol")
        base_clause = ", ".join(base_names) if base_names else "Protocol"
        lines.append("@runtime_checkable")
        lines.append(f"class {defn.name}({base_clause}):")
        if not defn.methods and not defn.attributes:
            lines.append("    ...")
        else:
            for attr in defn.attributes:
                lines.append(f"    {attr.name}: {attr.type_name}")
            for method in defn.methods:
                params = ", ".join([f"{n}: {t}" for t, n in method.params])
                params = f"self, {params}" if params else "self"
                lines.append(f"    def {method.name}({params}) -> {method.return_type}: ...")
    elif defn.kind == "struct":
        lines.append(f"class {defn.name}(TypedDict, total=False):")
        if not defn.struct_fields:
            lines.append("    pass")
        else:
            for field in defn.struct_fields:
                lines.append(f"    {field.name}: {field.type_name}")
    elif defn.kind == "enum":
        lines.append(f"class {defn.name}(IntEnum):")
        if not defn.enum_members:
            lines.append("    pass")
        else:
            for member in defn.enum_members:
                # If explicit value provided, keep it; else use auto()
                if "=" in member:
                    lines.append(f"    {member}")
                else:
                    lines.append(f"    {member} = auto()")
    return "\n".join(lines)


def ensure_init_files(registry: dict[Path, list[str]], out_root: Path) -> None:
    """Generate __init__.pyi that re-export generated symbols per package."""
    for directory, symbols in registry.items():
        init_path = directory / "__init__.pyi"
        lines: list[str] = []
        for symbol in sorted(set(symbols)):
            lines.append(f"from .{symbol} import {symbol}")
        if symbols:
            all_list = ", ".join([f'\"{s}\"' for s in sorted(set(symbols))])
            lines.append(f"__all__ = [{all_list}]")
        content = "\n".join(lines) + "\n"
        init_path.write_text(content, encoding="utf-8")


def ensure_package_markers(out_root: Path) -> None:
    """Ensure every package directory under out_root has an __init__.pyi marker."""
    for dirpath in [out_root, *out_root.rglob("*")]:
        if dirpath.is_dir():
            init_path = dirpath / "__init__.pyi"
            if not init_path.exists():
                init_path.write_text("", encoding="utf-8")


def write_stub(defs: list[Definition], module_parts: list[str], out_root: Path, registry: dict[Path, list[str]]) -> None:
    if not defs:
        return
    mod_name = module_parts[-1]
    out_dir = out_root.joinpath(*module_parts[:-1])
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / f"{mod_name}.pyi"

    imports: list[str] = ["from __future__ import annotations"]
    needs_any = True
    needs_protocol = any(d.kind in {"interface", "service"} for d in defs)
    needs_runtime = needs_protocol
    needs_typed_dict = any(d.kind == "struct" for d in defs)
    needs_intenum = any(d.kind == "enum" for d in defs)
    needs_auto = any(d.kind == "enum" and d.enum_members for d in defs)

    typing_imports: list[str] = []
    if needs_any:
        typing_imports.append("Any")
    if needs_protocol:
        typing_imports.append("Protocol")
    if needs_typed_dict:
        typing_imports.append("TypedDict")
    if needs_runtime:
        typing_imports.append("runtime_checkable")
    if typing_imports:
        imports.append(f"from typing import {', '.join(sorted(typing_imports))}")
    if needs_intenum:
        imports.append("from enum import IntEnum")
    if needs_auto:
        imports.append("from enum import auto")

    # Parent imports
    parent_imps: list[str] = []
    for d in defs:
        parent_imps.extend(parent_imports(d.parents))
    if parent_imps:
        imports.extend(sorted(set(parent_imps)))

    body_parts: list[str] = []
    for d in defs:
        body_parts.append(render_definition(d))
        body_parts.append("")

    content = "\n".join(imports) + "\n\n" + "\n".join(body_parts)
    out_file.write_text(content.strip() + "\n", encoding="utf-8")

    # Register for __init__ generation
    registry.setdefault(out_dir, []).append(mod_name)


def walk_idl_files(idl_root: Path) -> Iterable[Path]:
    return sorted(idl_root.rglob("*.idl"))


def module_parts_from_path(path: Path, idl_root: Path) -> list[str]:
    rel = path.relative_to(idl_root)
    parts = list(rel.with_suffix("").parts)
    # idl_root points to offapi/com, so rel starts with sun/star/...; prefix explicit com
    return ["com", *parts]


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate .pyi stubs from LibreOffice IDL")
    parser.add_argument("--idl-root", type=Path, default=Path(__file__).resolve().parent.parent / "offapi" / "com")
    parser.add_argument("--out-root", type=Path, default=Path(__file__).resolve().parent.parent / "src" / "stubs")
    args = parser.parse_args()

    idl_root: Path = args.idl_root
    out_root: Path = args.out_root
    out_root.mkdir(parents=True, exist_ok=True)

    registry: dict[Path, list[str]] = {}
    files = list(walk_idl_files(idl_root))
    for file in files:
        text = file.read_text(encoding="utf-8", errors="ignore")
        defs = parse_definitions(text)
        if not defs:
            continue
        parts = module_parts_from_path(file, idl_root)
        write_stub(defs, parts, out_root, registry)

    ensure_init_files(registry, out_root)
    ensure_package_markers(out_root)
    print(f"Generated stubs for {len(registry)} packages under {out_root}")


if __name__ == "__main__":
    main()
