#!/usr/bin/env python3
"""Split rbs_wbs_eventonight.html into two HTML files:
- rbs_eventonight.html: keeps only tags with class v*
- wbs_eventonight.html: keeps only tags with class a*

Usage:
    python split_rbs_wbs.py [input.html] [--rbs-output FILE] [--wbs-output FILE]
"""

from __future__ import annotations

import argparse
import html
from html.parser import HTMLParser
from pathlib import Path
from typing import List, Optional, Tuple, Union

VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input",
    "link", "meta", "param", "source", "track", "wbr",
}

PRESERVE_TAGS = {
    "html", "head", "body", "title", "style", "h1", "p", "div",
}


class TextNode:
    def __init__(self, text: str):
        self.text = text


class ElementNode:
    def __init__(self, tag: str, attrs: List[Tuple[str, Optional[str]]]):
        self.tag = tag
        self.attrs = attrs
        self.children: List[Node] = []


Node = Union[TextNode, ElementNode]


class HtmlTreeBuilder(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=False)
        self.root = ElementNode("document", [])
        self.stack: List[ElementNode] = [self.root]

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        node = ElementNode(tag.lower(), [(k.lower(), v) for k, v in attrs])
        self.stack[-1].children.append(node)
        if tag.lower() not in VOID_TAGS:
            self.stack.append(node)

    def handle_startendtag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        node = ElementNode(tag.lower(), [(k.lower(), v) for k, v in attrs])
        self.stack[-1].children.append(node)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                break

    def handle_data(self, data: str) -> None:
        if data.strip():
            self.stack[-1].children.append(TextNode(data.strip()))

    def handle_comment(self, data: str) -> None:
        # Commenti non necessari per i file derivati.
        pass


def class_matches(node: ElementNode, prefix: str) -> bool:
    if node.tag != "span":
        return False
    cls = None
    for key, value in node.attrs:
        if key == "class":
            cls = value or ""
            break
    if not cls:
        return False
    parts = cls.split()
    return any(part.startswith(prefix) for part in parts)


def keep_tree(node: Node, prefix: str) -> Optional[Node]:
    if isinstance(node, TextNode):
        return node if node.text.strip() else None

    if node.tag == "span" and not class_matches(node, prefix):
        return None

    kept_children: List[Node] = []
    for child in node.children:
        kept = keep_tree(child, prefix)
        if kept is None:
            continue
        if isinstance(kept, list):
            kept_children.extend(kept)
        else:
            kept_children.append(kept)

    if node.tag in {"ul", "li"} and not kept_children:
        return None

    if node.tag == "document":
        return kept_children

    if node.tag in VOID_TAGS:
        return node

    if node.tag in PRESERVE_TAGS or kept_children:
        new_node = ElementNode(node.tag, node.attrs)
        new_node.children = kept_children
        return new_node

    return None


def render(node: Node, indent: int = 0) -> str:
    space = "  " * indent
    if isinstance(node, TextNode):
        return f"{space}{html.escape(node.text, quote=False)}"

    attrs = ""
    if node.attrs:
        attrs = " " + " ".join(
            f'{name}="{html.escape(value, quote=True)}"' if value is not None else name
            for name, value in node.attrs
        )

    if node.tag in VOID_TAGS:
        return f"{space}<{node.tag}{attrs}>"

    if not node.children:
        return f"{space}<{node.tag}{attrs}></{node.tag}>"

    if len(node.children) == 1 and isinstance(node.children[0], TextNode):
        return f"{space}<{node.tag}{attrs}>{html.escape(node.children[0].text, quote=False)}</{node.tag}>"

    lines = [f"{space}<{node.tag}{attrs}>"]
    for child in node.children:
        lines.append(render(child, indent + 1))
    lines.append(f"{space}</{node.tag}>")
    return "\n".join(lines)


def replace_title(nodes: List[Node], title_text: str) -> None:
    def visit(node: Node) -> None:
        if isinstance(node, ElementNode):
            if node.tag == "title":
                node.children = [TextNode(title_text)]
                return
            for child in node.children:
                visit(child)

    for node in nodes:
        visit(node)


def split_html(input_path: Path, rbs_output: Path, wbs_output: Path) -> None:
    source = input_path.read_text(encoding="utf-8")
    parser = HtmlTreeBuilder()
    parser.feed(source)

    rbs_tree = keep_tree(parser.root, "v")
    wbs_tree = keep_tree(parser.root, "a")

    if not isinstance(rbs_tree, list) or not isinstance(wbs_tree, list):
        raise RuntimeError("Parsing interno fallito: albero HTML non valido.")

    replace_title(rbs_tree, "RBS – Espansione EvenToNight")
    replace_title(wbs_tree, "WBS – Espansione EvenToNight")

    def assemble(children: List[Node]) -> str:
        rendered = [render(child, 0) for child in children]
        return "\n".join(["<!DOCTYPE html>", *rendered, ""])

    rbs_output.write_text(assemble(rbs_tree), encoding="utf-8")
    wbs_output.write_text(assemble(wbs_tree), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Split del file HTML in una versione RBS e una WBS.")
    parser.add_argument(
        "input",
        nargs="?",
        default="rbs_wbs_eventonight.html",
        help="File HTML di partenza (default: rbs_wbs_eventonight.html)",
    )
    parser.add_argument(
        "--rbs-output",
        default="rbs_eventonight.html",
        help="File di output per la parte RBS (default: rbs_eventonight.html)",
    )
    parser.add_argument(
        "--wbs-output",
        default="wbs_eventonight.html",
        help="File di output per la parte WBS (default: wbs_eventonight.html)",
    )
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    base_dir = input_path.parent

    def resolve_output(output_arg: str) -> Path:
        output_path = Path(output_arg).expanduser()
        if output_path.is_absolute():
            return output_path.resolve()
        return (base_dir / output_path).resolve()

    rbs_output = resolve_output(args.rbs_output)
    wbs_output = resolve_output(args.wbs_output)

    split_html(input_path, rbs_output, wbs_output)
    print(f"Creati: {rbs_output.name}, {wbs_output.name}")


if __name__ == "__main__":
    main()
