#!/usr/bin/env python3
from __future__ import annotations

import re
import shutil
import os
from pathlib import Path

import yaml


ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "../open-xml-docs/docs").resolve()
SRC = ROOT / "src"


INCLUDE_RE = re.compile(r"\[!include\[[^\]]*\]\(([^)]+)\)\]")
CODE_RE = re.compile(r"\[!code-([A-Za-z0-9_+-]+)\[\]\(([^)#]+)(?:#([^)]+))?\)\]")
XREF_RE = re.compile(r"<xref:([^>]+)>")
BROKEN_XREF_RE = re.compile(r"(?<!<)xref:([A-Za-z0-9_.]+)\*?>")
FRONT_MATTER_RE = re.compile(r"\A---\n.*?\n---\n", re.DOTALL)
ROOT_LINK_RE = re.compile(r"(\[[^\]]+\]\()/(?!/)([^)\s]+)(\))")
LOCAL_LINK_RE = re.compile(r"(!?\[[^\]]*\]\()((?!https?://|mailto:|#|/)[^)\s]+)((?:\s+\"[^\"]*\")?\))")


def strip_front_matter(text: str) -> str:
    return FRONT_MATTER_RE.sub("", text, count=1).lstrip()


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeError:
            pass
    return path.read_text(encoding="latin-1")


def expand_includes(text: str, current_file: Path, stack: tuple[Path, ...] = ()) -> str:
    def replace(match: re.Match[str]) -> str:
        include_path = match.group(1).split("#", 1)[0]
        resolved = (current_file.parent / include_path).resolve()
        if resolved in stack:
            raise RuntimeError(f"recursive include: {resolved}")
        included = read_text(resolved)
        included = strip_front_matter(included)
        included = expand_includes(included.strip(), resolved, stack + (resolved,))
        return rewrite_local_links(included, resolved, current_file)

    return INCLUDE_RE.sub(replace, text)


def rewrite_local_links(text: str, from_file: Path, to_file: Path) -> str:
    def replace(match: re.Match[str]) -> str:
        prefix, target, suffix = match.groups()
        path, anchor = target.split("#", 1) if "#" in target else (target, "")
        if not path:
            return match.group(0)
        resolved = (from_file.parent / path).resolve()
        relative = os.path.relpath(resolved, to_file.parent).replace(os.sep, "/")
        if anchor:
            relative = f"{relative}#{anchor}"
        return f"{prefix}{relative}{suffix}"

    return LOCAL_LINK_RE.sub(replace, text)


def extract_snippet(source_file: Path, snippet: str | None) -> str:
    text = read_text(source_file)
    if snippet is None:
        selected = text
    else:
        requested = snippet
        if snippet.lower().startswith("extsnippet"):
            snippet = "Snippet0"
        start_re = re.compile(rf"<{re.escape(snippet)}\s*>", re.IGNORECASE)
        end_re = re.compile(rf"</{re.escape(snippet)}\s*>?", re.IGNORECASE)
        lines = text.splitlines()
        start = None
        depth = 0
        end = None
        for index, line in enumerate(lines):
            if start_re.search(line):
                if start is None:
                    start = index + 1
                depth += 1
                continue
            if start is not None and end_re.search(line):
                depth -= 1
                if depth == 0:
                    end = index
                    break
        if start is None or end is None:
            raise RuntimeError(f"snippet {requested} not found in {source_file}")
        selected = "\n".join(lines[start:end])

    selected = re.sub(r"(?m)^\s*(//|')\s*</?Snippet[^>]*>?\s*$\n?", "", selected)
    selected = selected.strip("\n")
    return selected


def expand_code_blocks(text: str, current_file: Path) -> str:
    def replace(match: re.Match[str]) -> str:
        language = match.group(1)
        sample_ref = match.group(2)
        snippet = match.group(3)
        if "?name=" in sample_ref:
            sample_ref, query_snippet = sample_ref.split("?name=", 1)
            snippet = snippet or query_snippet
        sample_ref = sample_ref.rstrip("?")
        sample_file = (current_file.parent / sample_ref).resolve()
        code = extract_snippet(sample_file, snippet)
        return f"```{language}\n{code}\n```"

    return CODE_RE.sub(replace, text)


def normalize_markdown(text: str) -> str:
    text = text.replace(r"\/", "/")
    text = XREF_RE.sub(lambda m: f"`{m.group(1).rstrip('*')}`", text)
    text = BROKEN_XREF_RE.sub(lambda m: f"`{m.group(1).rstrip('*')}`", text)
    text = ROOT_LINK_RE.sub(lambda m: f"{m.group(1)}https://learn.microsoft.com/{m.group(2)}{m.group(3)}", text)
    text = re.sub(
        r"(?m)^> \[!(NOTE|IMPORTANT|TIP|WARNING|CAUTION)\]\s*$",
        lambda m: f"> **{m.group(1).title()}**",
        text,
    )
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.rstrip() + "\n"


def convert_markdown(source_file: Path, dest_file: Path) -> None:
    text = source_file.read_text(encoding="utf-8")
    text = strip_front_matter(text)
    text = expand_includes(text, source_file)
    text = expand_code_blocks(text, source_file)
    text = normalize_markdown(text)
    dest_file.parent.mkdir(parents=True, exist_ok=True)
    dest_file.write_text(text, encoding="utf-8")


def toc_items() -> list[dict]:
    toc = yaml.safe_load((SOURCE / "toc.yml").read_text(encoding="utf-8"))
    if not toc or toc[0].get("name") != "Open XML SDK":
        raise RuntimeError("unexpected docs/toc.yml shape")
    return toc[0]["items"]


def collect_hrefs(items: list[dict]) -> list[str]:
    hrefs: list[str] = []
    seen: set[str] = set()

    def walk(nodes: list[dict]) -> None:
        for node in nodes:
            href = node.get("href")
            if href and href not in seen:
                seen.add(href)
                hrefs.append(href)
            if "items" in node:
                walk(node["items"])

    walk(items)
    return hrefs


def summary_link(title: str, href: str, indent: int) -> str:
    return f"{'  ' * indent}- [{title}]({href})"


def build_summary(items: list[dict]) -> str:
    lines = ["# Summary", "", "[Preface](preface.md)", ""]

    for node in items:
        if "items" not in node:
            lines.append(summary_link(node["name"], node["href"], 0))
            continue

        children = node["items"]
        parent_href = node.get("href")
        child_start = 0
        if not parent_href and children and children[0].get("name") == "Overview":
            parent_href = children[0].get("href")
            child_start = 1

        if parent_href:
            lines.append(summary_link(node["name"], parent_href, 0))
        else:
            lines.append(f"- {node['name']}")

        for child in children[child_start:]:
            if "href" in child:
                lines.append(summary_link(child["name"], child["href"], 1))

    return "\n".join(lines).rstrip() + "\n"


def write_preface() -> None:
    text = """# Preface

This mdBook starts from a baseline import of the Microsoft Open XML SDK documentation from the public `OfficeDev/open-xml-docs` repository.

The imported baseline content remains Microsoft documentation and is included only as source material for later ooxmlsdk-focused rewriting, including replacing the C# examples with Rust examples. Pages that still substantially preserve the Microsoft text, structure, examples, or images should be treated as derived from the upstream Microsoft documentation, not as original ooxmlsdk documentation.

New ooxmlsdk-specific writing, Rust examples, tooling, and repository-maintained material are licensed under this repository's `LICENSE-MIT` or `LICENSE-APACHE` files, unless a file states otherwise. Those licenses do not override third-party rights in the imported Microsoft baseline content.

Some imported articles quote or adapt text from ISO/IEC 29500. Those sections retain the ISO/IEC 29500 notices carried by the upstream documentation.
"""
    (SRC / "preface.md").write_text(text, encoding="utf-8")


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"source docs not found: {SOURCE}")

    for path in [
        SRC / "general",
        SRC / "presentation",
        SRC / "spreadsheet",
        SRC / "word",
        SRC / "migration",
        SRC / "media",
        SRC / "includes",
    ]:
        if path.exists():
            shutil.rmtree(path)

    for path in [SRC / "chapter_1.md", SRC / "preface.md", SRC / "SUMMARY.md"]:
        if path.exists():
            path.unlink()

    if (SOURCE / "media").exists():
        shutil.copytree(SOURCE / "media", SRC / "media")

    items = toc_items()
    for href in collect_hrefs(items):
        convert_markdown(SOURCE / href, SRC / href)

    write_preface()
    (SRC / "SUMMARY.md").write_text(build_summary(items), encoding="utf-8")


if __name__ == "__main__":
    main()
