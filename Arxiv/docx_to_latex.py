#!/usr/bin/env python3
"""
Convert a DOCX document into a LaTeX .tex file without losing textual content.

This script avoids third-party dependencies so that it can run in constrained
environments. It parses the OpenXML structure directly and emits a LaTeX
document with a reasonable structure, preserving headings, lists, tables,
figures, and inline formatting where practical.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {
    "w": W_NS,
    "m": M_NS,
    "r": R_NS,
    "a": A_NS,
    "wp": WP_NS,
    "pic": PIC_NS,
    "rel": REL_NS,
}

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
EMU_PER_INCH = 914400  # DrawingML measurement units


def escape_latex(text: str) -> str:
    """Escape special LaTeX characters in a text fragment."""
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(ch, ch) for ch in text)


def normalize_whitespace(text: str) -> str:
    """Collapse excessive whitespace but preserve single spaces."""
    # Replace Windows newlines etc. with single spaces so paragraphs flow.
    collapsed = re.sub(r"\s+", " ", text)
    return collapsed.strip()


def tex_path(path: Path, base_dir: Path) -> str:
    """Return a POSIX-style path relative to base_dir for LaTeX include."""
    rel_path = path.relative_to(base_dir)
    return rel_path.as_posix()


@dataclass
class ListState:
    level: int
    env: str
    num_id: str


class DocxToLatexConverter:
    def __init__(self, docx_path: Path, tex_path: Path) -> None:
        self.docx_path = docx_path
        self.tex_path = tex_path
        self.output_dir = tex_path.parent
        self.media_dir = self.output_dir / f"{tex_path.stem}_media"
        self.zip = zipfile.ZipFile(docx_path)
        self.relationships = self._load_relationships()
        self.num_formats = self._load_numbering()
        self.list_stack: List[ListState] = []
        self.footnotes = self._load_notes("word/footnotes.xml")
        self.endnotes = self._load_notes("word/endnotes.xml")

    def convert(self) -> None:
        """Perform the conversion and write the LaTeX output."""
        document_xml = self.zip.read("word/document.xml")
        doc_root = ET.fromstring(document_xml)
        body = doc_root.find("w:body", NS)
        if body is None:
            raise ValueError("DOCX body not found")

        self.media_dir.mkdir(parents=True, exist_ok=True)
        latex_lines: List[str] = []
        latex_lines.extend(self._latex_preamble())
        latex_lines.extend(self._convert_headers(doc_root))
        latex_lines.append("")  # ensure a blank line after header content

        for child in body:
            latex_lines.extend(self._convert_block(child))

        latex_lines.extend(self._close_all_lists())
        latex_lines.append("\\end{document}")

        self.tex_path.write_text("\n".join(self._clean_lines(latex_lines)), encoding="utf-8")

    def _latex_preamble(self) -> List[str]:
        return [
            "\\documentclass[11pt]{article}",
            "\\usepackage[utf8]{inputenc}",
            "\\usepackage[T1]{fontenc}",
            "\\usepackage{geometry}",
            "\\usepackage{array}",
            "\\usepackage{longtable}",
            "\\usepackage{graphicx}",
            "\\usepackage{amsmath}",
            "\\usepackage{amssymb}",
            "\\usepackage{hyperref}",
            "\\usepackage{enumitem}",
            "\\usepackage{multirow}",
            "\\geometry{margin=1in}",
            "\\setlist[itemize]{itemsep=0pt, topsep=4pt}",
            "\\setlist[enumerate]{itemsep=0pt, topsep=4pt}",
            "\\begin{document}",
        ]

    def _convert_headers(self, doc_root: ET.Element) -> List[str]:
        """Convert any header references attached to the document."""
        result: List[str] = []
        sect_pr = doc_root.find(".//w:sectPr", NS)
        if sect_pr is None:
            return result
        for header_ref in sect_pr.findall("w:headerReference", NS):
            rel_id = header_ref.attrib.get(f"{{{R_NS}}}id")
            if not rel_id:
                continue
            rel_info = self.relationships.get(rel_id)
            if not rel_info:
                continue
            _, target = rel_info
            header_path = f"word/{target}"
            if not self._zip_has(header_path):
                continue
            header_root = ET.fromstring(self.zip.read(header_path))
            for child in header_root.findall("w:p", NS):
                paragraph_lines = self._convert_paragraph(child, in_table=False, suppress_list=True)
                result.extend(paragraph_lines)
        return result

    def _zip_has(self, member: str) -> bool:
        try:
            self.zip.getinfo(member)
            return True
        except KeyError:
            return False

    def _load_relationships(self) -> Dict[str, Tuple[str, str]]:
        rels_path = "word/_rels/document.xml.rels"
        relationships: Dict[str, Tuple[str, str]] = {}
        if not self._zip_has(rels_path):
            return relationships
        rels_root = ET.fromstring(self.zip.read(rels_path))
        for rel in rels_root:
            rel_id = rel.attrib.get("Id")
            rel_type = rel.attrib.get("Type")
            target = rel.attrib.get("Target")
            if rel_id and rel_type and target:
                relationships[rel_id] = (rel_type, target)
        return relationships

    def _load_numbering(self) -> Dict[str, Dict[int, str]]:
        numbering_path = "word/numbering.xml"
        if not self._zip_has(numbering_path):
            return {}
        root = ET.fromstring(self.zip.read(numbering_path))
        abstract_map: Dict[str, Dict[int, str]] = {}
        for abstract in root.findall("w:abstractNum", NS):
            abstract_id = abstract.attrib.get(f"{{{W_NS}}}abstractNumId")
            if not abstract_id:
                continue
            levels: Dict[int, str] = {}
            for lvl in abstract.findall("w:lvl", NS):
                ilvl = lvl.attrib.get(f"{{{W_NS}}}ilvl")
                numfmt = lvl.find("w:numFmt", NS)
                if ilvl is None or numfmt is None:
                    continue
                fmt = numfmt.attrib.get(f"{{{W_NS}}}val")
                if fmt:
                    levels[int(ilvl)] = fmt
            abstract_map[abstract_id] = levels

        num_map: Dict[str, Dict[int, str]] = defaultdict(dict)
        for num in root.findall("w:num", NS):
            num_id = num.attrib.get(f"{{{W_NS}}}numId")
            abstract_ref = num.find("w:abstractNumId", NS)
            if not num_id or abstract_ref is None:
                continue
            abstract_id = abstract_ref.attrib.get(f"{{{W_NS}}}val")
            if not abstract_id:
                continue
            levels = abstract_map.get(abstract_id, {})
            num_map[num_id] = levels
        return num_map

    def _load_notes(self, path: str) -> Dict[str, List[str]]:
        if not self._zip_has(path):
            return {}
        root = ET.fromstring(self.zip.read(path))
        notes: Dict[str, List[str]] = {}
        for note in root.findall("w:footnote", NS) + root.findall("w:endnote", NS):
            note_id = note.attrib.get(f"{{{W_NS}}}id")
            if not note_id:
                continue
            parts: List[str] = []
            for child in note:
                parts.extend(self._convert_block(child))
            notes[note_id] = parts
        return notes

    def _convert_block(self, element: ET.Element) -> List[str]:
        tag = element.tag
        if tag == self._q("w:p"):
            return self._convert_paragraph(element)
        if tag == self._q("w:tbl"):
            lines = []
            lines.extend(self._close_all_lists())
            lines.extend(self._convert_table(element))
            lines.append("")
            return lines
        if tag == self._q("w:sectPr"):
            return self._close_all_lists()
        # Unsupported block-level element; skip but ensure lists closed.
        return []

    def _convert_paragraph(
        self,
        paragraph: ET.Element,
        *,
        in_table: bool = False,
        suppress_list: bool = False,
    ) -> List[str]:
        list_info = None if suppress_list else self._extract_list_info(paragraph)
        lines: List[str] = []
        if not in_table:
            lines.extend(self._sync_list_stack(list_info))

        text_content, block_fragments = self._collect_runs(paragraph)
        text_content = normalize_whitespace(text_content)

        style = self._paragraph_style(paragraph)
        if list_info and not suppress_list:
            if not text_content and not block_fragments:
                return lines
            item_line = "\\item"
            if text_content:
                item_line += f" {text_content}"
            lines.append(item_line)
            lines.extend(block_fragments)
            return lines

        # Not inside a list
        if style == "Heading1" and text_content:
            lines.append(f"\\section{{{text_content}}}")
        elif style == "Heading2" and text_content:
            lines.append(f"\\subsection{{{text_content}}}")
        elif style == "Heading3" and text_content:
            lines.append(f"\\subsubsection{{{text_content}}}")
        elif style == "Title" and text_content:
            lines.append(f"\\title{{{text_content}}}")
            lines.append("\\maketitle")
        elif text_content:
            lines.append(text_content)
        lines.extend(block_fragments)
        if lines and not lines[-1]:
            return lines
        if not in_table and lines:
            lines.append("")  # blank line between paragraphs
        return lines

    def _collect_runs(self, paragraph: ET.Element) -> Tuple[str, List[str]]:
        fragments: List[str] = []
        blocks: List[str] = []
        for child in paragraph:
            if child.tag == self._q("w:r"):
                text, block = self._convert_run(child)
                fragments.append(text)
                blocks.extend(block)
            elif child.tag == self._q("w:hyperlink"):
                text, block = self._convert_hyperlink(child)
                fragments.append(text)
                blocks.extend(block)
            elif child.tag == self._q("w:fldSimple"):
                # Field codes - render their simple text content.
                text_parts = []
                for run in child.findall("w:r", NS):
                    text, block = self._convert_run(run)
                    text_parts.append(text)
                    blocks.extend(block)
                fragments.append("".join(text_parts))
            elif child.tag in (self._q("m:oMath"), self._q("m:oMathPara")):
                math_text = self._convert_math(child)
                fragments.append(math_text)
            elif child.tag == self._q("w:proofErr") or child.tag == self._q("w:bookmarkStart") or child.tag == self._q("w:bookmarkEnd"):
                continue
        return "".join(fragments), blocks

    def _convert_hyperlink(self, hyperlink: ET.Element) -> Tuple[str, List[str]]:
        rel_id = hyperlink.attrib.get(f"{{{R_NS}}}id")
        inner_text_parts: List[str] = []
        blocks: List[str] = []
        for run in hyperlink.findall("w:r", NS):
            text, new_blocks = self._convert_run(run)
            inner_text_parts.append(text)
            blocks.extend(new_blocks)
        text_value = "".join(inner_text_parts)
        if rel_id and rel_id in self.relationships:
            _, target = self.relationships[rel_id]
            href = escape_latex(target)
            if text_value:
                return f"\\href{{{href}}}{{{text_value}}}", blocks
        return text_value, blocks

    def _convert_run(self, run: ET.Element) -> Tuple[str, List[str]]:
        blocks: List[str] = []
        texts: List[str] = []
        for child in run:
            if child.tag == self._q("w:t"):
                texts.append(self._format_run_text(child, run))
            elif child.tag == self._q("w:tab"):
                texts.append("\\hspace*{1em}")
            elif child.tag == self._q("w:br"):
                texts.append("\\\\")
            elif child.tag == self._q("w:drawing"):
                figure_block = self._convert_drawing(child)
                if figure_block:
                    blocks.extend(figure_block)
            elif child.tag == self._q("w:pict"):
                # Legacy picture; treat as unsupported block to avoid loss.
                blocks.append("% Picture element omitted (unsupported)")
            elif child.tag == self._q("m:oMath") or child.tag == self._q("m:oMathPara"):
                math_text = self._convert_math(child)
                if math_text:
                    texts.append(math_text)
        return "".join(texts), blocks

    def _format_run_text(self, text_node: ET.Element, run: ET.Element) -> str:
        raw_text = text_node.text or ""
        escaped = escape_latex(raw_text)
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            return escaped
        wrappers: List[str] = []
        if rpr.find("w:b", NS) is not None:
            wrappers.append("textbf")
        if rpr.find("w:i", NS) is not None:
            wrappers.append("emph")
        underline = rpr.find("w:u", NS)
        if underline is not None and underline.attrib.get(f"{{{W_NS}}}val", "").lower() != "none":
            wrappers.append("underline")
        vert_align = rpr.find("w:vertAlign", NS)
        if vert_align is not None:
            align = vert_align.attrib.get(f"{{{W_NS}}}val")
            if align == "superscript":
                escaped = f"\\textsuperscript{{{escaped}}}"
            elif align == "subscript":
                escaped = f"\\textsubscript{{{escaped}}}"
        for wrapper in wrappers:
            escaped = f"\\{wrapper}{{{escaped}}}"
        return escaped

    def _convert_math(self, element: ET.Element) -> str:
        # Collect all math text fragments; fallback to textual representation.
        parts: List[str] = []
        for node in element.findall(".//m:t", NS):
            if node.text:
                parts.append(node.text)
        if not parts:
            return ""
        math_text = " ".join(parts)
        return f"${math_text}$"

    def _convert_drawing(self, drawing: ET.Element) -> Optional[List[str]]:
        inline = drawing.find(".//wp:inline", NS)
        anchor = drawing.find(".//wp:anchor", NS) if inline is None else None
        container = inline if inline is not None else anchor
        if container is None:
            return None
        blip = container.find(".//a:blip", NS)
        if blip is None:
            return None
        embed_id = blip.attrib.get(f"{{{R_NS}}}embed")
        if not embed_id:
            return None
        rel_info = self.relationships.get(embed_id)
        if not rel_info:
            return None
        rel_type, target = rel_info
        if not rel_type.endswith("/image"):
            return None
        sanitized_target = target.replace("\\", "/")
        zip_member = f"word/{sanitized_target}"
        try:
            data = self.zip.read(zip_member)
        except KeyError:
            return None
        dest_path = self.media_dir / Path(target).name
        if not dest_path.exists():
            dest_path.write_bytes(data)
        width_option = "width=0.7\\linewidth"
        extent = container.find("wp:extent", NS)
        if extent is not None:
            cx = extent.attrib.get("cx")
            if cx:
                try:
                    width_in = int(cx) / EMU_PER_INCH
                    width_option = f"width={width_in:.2f}in"
                except ValueError:
                    pass
        doc_pr = container.find("wp:docPr", NS)
        caption: Optional[str] = None
        if doc_pr is not None:
            caption = doc_pr.attrib.get("descr") or doc_pr.attrib.get("title")
        latex_lines = [
            "\\begin{figure}[h]",
            "\\centering",
            f"\\includegraphics[{width_option}]{{{escape_latex(tex_path(dest_path, self.output_dir))}}}",
        ]
        if caption:
            latex_lines.append(f"\\caption{{{escape_latex(normalize_whitespace(caption))}}}")
        latex_lines.append("\\end{figure}")
        return latex_lines

    def _convert_table(self, table: ET.Element) -> List[str]:
        rows_data: List[List[Tuple[str, int]]] = []
        max_cols = 0
        for tr in table.findall("w:tr", NS):
            row_cells: List[Tuple[str, int]] = []
            for tc in tr.findall("w:tc", NS):
                span = 1
                cell_pr = tc.find("w:tcPr", NS)
                if cell_pr is not None:
                    grid_span = cell_pr.find("w:gridSpan", NS)
                    if grid_span is not None:
                        try:
                            span = int(grid_span.attrib.get(f"{{{W_NS}}}val", "1"))
                        except ValueError:
                            span = 1
                cell_text = self._convert_table_cell(tc)
                row_cells.append((cell_text, span))
            span_total = sum(span for _, span in row_cells) or 1
            max_cols = max(max_cols, span_total)
            rows_data.append(row_cells)

        if max_cols == 0:
            max_cols = 1
        column_spec = "|" + "|".join([f"p{{\\dimexpr\\linewidth/{max_cols}-2\\tabcolsep\\relax}}" for _ in range(max_cols)]) + "|"
        table_lines = [f"\\begin{{longtable}}{{{column_spec}}}", "\\hline"]
        for row in rows_data:
            line_parts: List[str] = []
            for text, span in row:
                cell_content = text if text else "\\quad{}"
                if span > 1:
                    line_parts.append(f"\\multicolumn{{{span}}}{{|p{{\\dimexpr\\linewidth/{max_cols}-2\\tabcolsep\\relax}}|}}{{{cell_content}}}")
                else:
                    line_parts.append(cell_content)
            table_lines.append(" & ".join(line_parts) + r" \\")
            table_lines.append("\\hline")
        table_lines.append("\\end{longtable}")
        return table_lines

    def _convert_table_cell(self, cell: ET.Element) -> str:
        parts: List[str] = []
        for paragraph in cell.findall("w:p", NS):
            text, blocks = self._collect_runs(paragraph)
            content = text.strip()
            if content:
                parts.append(content)
            parts.extend(blocks)
        return r" \\ ".join(part for part in parts if part)

    def _paragraph_style(self, paragraph: ET.Element) -> Optional[str]:
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            return None
        style = ppr.find("w:pStyle", NS)
        if style is None:
            return None
        return style.attrib.get(f"{{{W_NS}}}val")

    def _extract_list_info(self, paragraph: ET.Element) -> Optional[ListState]:
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            return None
        num_pr = ppr.find("w:numPr", NS)
        if num_pr is None:
            return None
        num_id_elem = num_pr.find("w:numId", NS)
        if num_id_elem is None:
            return None
        num_id = num_id_elem.attrib.get(f"{{{W_NS}}}val")
        if not num_id:
            return None
        ilvl_elem = num_pr.find("w:ilvl", NS)
        try:
            level = int(ilvl_elem.attrib.get(f"{{{W_NS}}}val", "0")) if ilvl_elem is not None else 0
        except ValueError:
            level = 0
        fmt = self.num_formats.get(num_id, {}).get(level, "bullet")
        env = "itemize" if fmt == "bullet" else "enumerate"
        return ListState(level=level, env=env, num_id=num_id)

    def _sync_list_stack(self, target: Optional[ListState]) -> List[str]:
        lines: List[str] = []
        if target is None:
            lines.extend(self._close_all_lists())
            return lines
        # Close lists deeper than the target level.
        while self.list_stack and self.list_stack[-1].level > target.level:
            state = self.list_stack.pop()
            lines.append(f"\\end{{{state.env}}}")
        # If same level but different environment, close and reopen.
        if self.list_stack and self.list_stack[-1].level == target.level:
            current = self.list_stack[-1]
            if current.env != target.env or current.num_id != target.num_id:
                state = self.list_stack.pop()
                lines.append(f"\\end{{{state.env}}}")
        # Open new levels as needed.
        while not self.list_stack or self.list_stack[-1].level < target.level:
            new_state = ListState(level=len(self.list_stack), env=target.env, num_id=target.num_id)
            self.list_stack.append(new_state)
            lines.append(f"\\begin{{{target.env}}}")
        if not self.list_stack or self.list_stack[-1].level < target.level:
            new_state = ListState(level=target.level, env=target.env, num_id=target.num_id)
            self.list_stack.append(new_state)
            lines.append(f"\\begin{{{target.env}}}")
        elif self.list_stack[-1].level == target.level and (self.list_stack[-1].env != target.env or self.list_stack[-1].num_id != target.num_id):
            state = self.list_stack.pop()
            lines.append(f"\\end{{{state.env}}}")
            self.list_stack.append(target)
            lines.append(f"\\begin{{{target.env}}}")
        elif not self.list_stack:
            self.list_stack.append(target)
            lines.append(f"\\begin{{{target.env}}}")
        return lines

    def _close_all_lists(self) -> List[str]:
        lines: List[str] = []
        while self.list_stack:
            state = self.list_stack.pop()
            lines.append(f"\\end{{{state.env}}}")
        return lines

    def _clean_lines(self, lines: Iterable[str]) -> List[str]:
        cleaned: List[str] = []
        previous_blank = False
        for line in lines:
            stripped = line.rstrip()
            if not stripped:
                if previous_blank:
                    continue
                previous_blank = True
                cleaned.append("")
                continue
            previous_blank = False
            cleaned.append(stripped)
        return cleaned

    def _q(self, tag: str) -> str:
        prefix, local = tag.split(":")
        return f"{{{NS[prefix]}}}{local}"


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert DOCX to LaTeX without third-party tools.")
    parser.add_argument("docx_path", type=Path, help="Path to the DOCX file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Path for the generated LaTeX file (defaults to DOCX name with .tex)",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(argv)
    docx_path: Path = args.docx_path
    if not docx_path.exists():
        raise SystemExit(f"Input file not found: {docx_path}")
    tex_path = args.output if args.output is not None else docx_path.with_suffix(".tex")
    tex_path.parent.mkdir(parents=True, exist_ok=True)

    converter = DocxToLatexConverter(docx_path, tex_path)
    converter.convert()
    print(f"Converted {docx_path} -> {tex_path}")


if __name__ == "__main__":
    main()
