"""
document_generator.py
Service for generating professional documents (DOCX and PDF) from user content.

This module provides AI-powered document generation that converts user queries
and content into formatted Word documents and PDFs.
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import fitz  # PyMuPDF for PDF generation

# Windows COM support for docx2pdf (optional)
try:
    if sys.platform == "win32":
        import pythoncom
except ImportError:
    pythoncom = None

from .docx_builder import (
    BulletBlock,
    CalloutBlock,
    CodeBlock,
    DocumentSpec,
    NumberedBlock,
    ParagraphBlock,
    SectionSpec,
    TableSpec,
    build_docx_bytes,
)


class DocumentGenerator:
    """Generate professional documents from user content."""

    @staticmethod
    def detect_document_request(user_message: str) -> tuple[bool, Optional[str]]:
        """
        Detect if user is requesting document generation.

        Returns:
            Tuple of (is_document_request, requested_format)
            Format is "docx", "pdf", or None (both formats)
        """
        message_lower = user_message.lower()

        # PDF-specific triggers
        pdf_triggers = [
            r"\bpdf\b",
            r"\bto\s+pdf\b",
            r"\bas\s+pdf\b",
        ]
        for pattern in pdf_triggers:
            if re.search(pattern, message_lower):
                return True, "pdf"

        # Document/report/word triggers — any occurrence is enough
        doc_keywords = r"\b(document|report|doc|docx|word\s+file|word\s+doc)\b"
        if re.search(doc_keywords, message_lower):
            return True, "docx"

        # Action verbs strongly implying a written output
        action_patterns = [
            r"\b(write|draft|prepare|produce|compose|generate|create|make|export|save|give\s+me)\b.{0,40}\b(summary|overview|guide|tutorial|notes|writeup|write.?up|analysis)\b",
        ]
        for pattern in action_patterns:
            if re.search(pattern, message_lower):
                return True, "docx"

        return False, None

    @staticmethod
    def parse_content_to_sections(content: str, title: str = "") -> DocumentSpec:
        """
        Parse user content and LLM response into structured document sections.

        Args:
            content: The main content to convert (can include markdown)
            title: Document title (extracted from content if not provided)

        Returns:
            DocumentSpec ready for document generation
        """
        sections = []

        # Extract title if not provided
        if not title:
            title_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
            if title_match:
                title = title_match.group(1).strip()
                # Remove the title line from content
                content = content.replace(title_match.group(0), "", 1).strip()
            else:
                title = "Generated Document"

        # Split content by markdown headers
        lines = content.split("\n")
        current_section = None
        current_blocks = []
        current_para = []

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # Heading level 1 (body)
            if re.match(r"^#(?!#) ", line):
                if current_section:
                    if current_para:
                        current_blocks.append(
                            ParagraphBlock(text=" ".join(current_para))
                        )
                        current_para = []
                    current_section.blocks = current_blocks
                    sections.append(current_section)

                heading = line[2:].strip()
                current_section = SectionSpec(heading=heading, level=1, blocks=[])
                current_blocks = []
                i += 1
                continue

            # Heading level 2
            elif line.startswith("## "):
                # Save previous section
                if current_section:
                    if current_para:
                        current_blocks.append(
                            ParagraphBlock(text=" ".join(current_para))
                        )
                        current_para = []
                    current_section.blocks = current_blocks
                    sections.append(current_section)

                # Start new section
                heading = line[3:].strip()
                current_section = SectionSpec(heading=heading, level=2, blocks=[])
                current_blocks = []
                i += 1
                continue

            # Heading level 3
            elif line.startswith("### "):
                # Save previous section
                if current_section:
                    if current_para:
                        current_blocks.append(
                            ParagraphBlock(text=" ".join(current_para))
                        )
                        current_para = []
                    current_section.blocks = current_blocks
                    sections.append(current_section)

                # Start new section
                heading = line[4:].strip()
                current_section = SectionSpec(heading=heading, level=3, blocks=[])
                current_blocks = []
                i += 1
                continue

            # Bullet list
            elif line.startswith(("- ", "* ", "• ")):
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []

                # Collect all bullet items
                items = []
                while i < len(lines) and lines[i].strip().startswith(("- ", "* ", "• ")):
                    item_text = lines[i].strip()[2:].strip()
                    items.append(item_text)
                    i += 1
                current_blocks.append(BulletBlock(items=items))
                continue

            # Numbered list
            elif re.match(r"^\d+[\.|\)]\s+", line):
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []

                # Collect all numbered items
                items = []
                while i < len(lines) and re.match(r"^\d+[\.|\)]\s+", lines[i].strip()):
                    item_text = re.sub(r"^\d+[\.|\)]\s+", "", lines[i].strip())
                    items.append(item_text)
                    i += 1
                current_blocks.append(NumberedBlock(items=items))
                continue

            # Code block
            elif line.startswith("```"):
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []

                # Extract language if specified
                language = line[3:].strip() if len(line) > 3 else ""
                i += 1

                # Collect code lines
                code_lines = []
                while i < len(lines) and not lines[i].strip().startswith("```"):
                    code_lines.append(lines[i])
                    i += 1

                code = "\n".join(code_lines)
                current_blocks.append(CodeBlock(code=code, language=language))
                i += 1  # Skip closing ```
                continue

            # Callout/Alert blocks
            elif line.startswith("> "):
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []

                # Collect callout content
                callout_lines = []
                while i < len(lines) and lines[i].strip().startswith("> "):
                    callout_lines.append(lines[i].strip()[2:])
                    i += 1

                callout_text = " ".join(callout_lines)

                # Detect callout type from prefix keyword (> WARNING: ...) or body
                prefix_match = re.match(r"^(warning|caution|danger|tip|hint|note|important|critical)\b[:\s]?", callout_text, re.IGNORECASE)
                callout_kind = "note"
                if prefix_match:
                    kw = prefix_match.group(1).lower()
                    if kw in ("warning", "caution", "danger"):
                        callout_kind = "warning"
                    elif kw in ("tip", "hint"):
                        callout_kind = "tip"
                    elif kw in ("important", "critical"):
                        callout_kind = "important"
                    # Strip the prefix from the text so it's not doubled
                    callout_text = callout_text[prefix_match.end():].strip()
                elif any(kw in callout_text.lower() for kw in ["warning", "caution", "danger"]):
                    callout_kind = "warning"
                elif any(kw in callout_text.lower() for kw in ["tip", "hint"]):
                    callout_kind = "tip"
                elif any(kw in callout_text.lower() for kw in ["important", "critical"]):
                    callout_kind = "important"

                current_blocks.append(
                    CalloutBlock(text=callout_text, kind=callout_kind)
                )
                continue

            # Horizontal rule — skip (used as section divider by some LLMs)
            elif re.match(r"^[-*_]{3,}$", line):
                i += 1
                continue

            # Standalone bold line → implicit heading level 3
            # e.g. **Section Name** or **Section Name:**
            elif re.match(r"^\*\*[^*]+\*\*:?$", line):
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []
                # Only treat as heading when not inside a section yet, or flush current section
                heading_text = re.sub(r"^\*\*(.+?)\*\*:?$", r"\1", line).strip()
                if current_section:
                    current_section.blocks = current_blocks
                    sections.append(current_section)
                current_section = SectionSpec(heading=heading_text, level=3, blocks=[])
                current_blocks = []
                i += 1
                continue

            # Markdown table  (lines that start with "|")  
            elif line.startswith("|"):
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []

                # Collect all consecutive table lines
                table_lines: list[str] = []
                while i < len(lines) and lines[i].strip().startswith("|"):
                    table_lines.append(lines[i].strip())
                    i += 1

                if len(table_lines) >= 2:
                    def _parse_row(tl: str) -> list[str]:
                        """Split a markdown table row into stripped cells."""
                        cells = [c.strip() for c in tl.split("|")]
                        # Remove leading / trailing empty strings from `|...|` syntax
                        if cells and cells[0] == "":
                            cells = cells[1:]
                        if cells and cells[-1] == "":
                            cells = cells[:-1]
                        return cells

                    headers = _parse_row(table_lines[0])
                    data_rows: list[list[str]] = []
                    for tl in table_lines[1:]:
                        cells = _parse_row(tl)
                        # Skip separator rows like |---|:---:|---|
                        if all(re.match(r"^[-: ]+$", c) for c in cells):
                            continue
                        data_rows.append(cells)

                    if headers:
                        current_blocks.append(
                            TableSpec(headers=headers, rows=data_rows)
                        )
                continue

            # Empty line
            elif not line:
                if current_para:
                    current_blocks.append(ParagraphBlock(text=" ".join(current_para)))
                    current_para = []
                i += 1
                continue

            # Regular paragraph text
            else:
                current_para.append(line)
                i += 1

        # Add remaining paragraph
        if current_para:
            current_blocks.append(ParagraphBlock(text=" ".join(current_para)))

        # Add final section
        if current_section:
            current_section.blocks = current_blocks
            sections.append(current_section)
        elif current_blocks:
            # No sections found, create a default one
            sections.append(
                SectionSpec(heading="Content", level=2, blocks=current_blocks)
            )

        # Create document spec
        doc_spec = DocumentSpec(
            title=title,
            author="ADI TestAgent",
            date=datetime.now().strftime("%B %d, %Y"),
            confidentiality_label="Internal Use Only",
            include_toc=len(sections) > 3,
            include_title_page=True,
            sections=sections,
        )

        return doc_spec

    @staticmethod
    def generate_docx(content: str, title: str = "") -> bytes:
        """
        Generate a Word document from content.

        Args:
            content: Main content (supports markdown)
            title: Document title

        Returns:
            DOCX file as bytes
        """
        doc_spec = DocumentGenerator.parse_content_to_sections(content, title)
        return build_docx_bytes(doc_spec)

    @staticmethod
    def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
        """
        Convert DOCX bytes to PDF bytes.

        Args:
            docx_bytes: DOCX file content as bytes

        Returns:
            PDF file as bytes
        """
        # Create temporary files
        with tempfile.NamedTemporaryFile(
            suffix=".docx", delete=False
        ) as temp_docx, tempfile.NamedTemporaryFile(
            suffix=".pdf", delete=False
        ) as temp_pdf:
            temp_docx_path = temp_docx.name
            temp_pdf_path = temp_pdf.name

            # Write DOCX
            temp_docx.write(docx_bytes)
            temp_docx.flush()

        try:
            # Convert DOCX to PDF using PyMuPDF
            # Note: PyMuPDF doesn't directly convert DOCX to PDF
            # We need to use an alternative approach

            # For Windows, we can use COM automation if available
            # For cross-platform, we'll use a different approach

            # Alternative: Use docx2pdf library if available
            try:
                import docx2pdf

                # Initialize COM for Windows (required for docx2pdf)
                if sys.platform == "win32" and pythoncom:
                    try:
                        pythoncom.CoInitialize()
                        docx2pdf.convert(temp_docx_path, temp_pdf_path)
                        pythoncom.CoUninitialize()
                    except Exception as e:
                        # If COM initialization fails, try without it
                        docx2pdf.convert(temp_docx_path, temp_pdf_path)
                else:
                    docx2pdf.convert(temp_docx_path, temp_pdf_path)
            except ImportError:
                # Fallback: Create a simple PDF with PyMuPDF from text
                # This is a basic fallback - not ideal but works
                doc = fitz.open()
                page = doc.new_page()

                # Read docx content and extract text
                from docx import Document as DocxDocument

                docx_doc = DocxDocument(temp_docx_path)
                text = "\n\n".join([para.text for para in docx_doc.paragraphs])

                # Insert text into PDF
                rect = page.rect
                # Leave margins
                text_rect = fitz.Rect(50, 50, rect.width - 50, rect.height - 50)
                page.insert_textbox(
                    text_rect, text, fontsize=11, fontname="helv", align=0
                )

                doc.save(temp_pdf_path)
                doc.close()

            # Read PDF bytes
            with open(temp_pdf_path, "rb") as f:
                pdf_bytes = f.read()

            return pdf_bytes

        finally:
            # Clean up temp files
            try:
                os.unlink(temp_docx_path)
            except Exception:
                pass
            try:
                os.unlink(temp_pdf_path)
            except Exception:
                pass

    @staticmethod
    def generate_pdf(content: str, title: str = "") -> bytes:
        """
        Generate a PDF document from content.

        Args:
            content: Main content (supports markdown)
            title: Document title

        Returns:
            PDF file as bytes
        """
        # First generate DOCX
        docx_bytes = DocumentGenerator.generate_docx(content, title)

        # Convert to PDF
        return DocumentGenerator.convert_docx_to_pdf(docx_bytes)

    @staticmethod
    def create_from_llm_response(
        user_query: str, llm_response: str, format: str = "docx"
    ) -> bytes:
        """
        Create a document from user query and LLM response.

        Args:
            user_query: Original user query
            llm_response: LLM's response content
            format: "docx" or "pdf"

        Returns:
            Document as bytes
        """
        # 1. Use the # heading from the LLM response if present (most reliable)
        h1_match = re.search(r"^#(?!#)\s+(.+)$", llm_response, re.MULTILINE)
        if h1_match:
            title = h1_match.group(1).strip()
            full_content = llm_response  # already has its own title heading
        else:
            # 2. Fall back: derive title from user query
            title = "Generated Report"
            # Try "about X", "on X", "for X" patterns
            for pattern in [
                r"\babout\s+(.+?)(?:\s*\.|$)",
                r"\bon\s+(.+?)(?:\s*\.|$)",
                r"\bfor\s+(.+?)(?:\s*\.|$)",
                r"\bregarding\s+(.+?)(?:\s*\.|$)",
            ]:
                m = re.search(pattern, user_query, re.IGNORECASE)
                if m:
                    title = m.group(1).strip().title()
                    break
            else:
                # Last resort: use the first 60 chars of the query
                title = user_query.strip()[:60].title()
                if len(user_query) > 60:
                    title = title.rstrip() + "..."

            full_content = f"# {title}\n\n{llm_response}"

        if format.lower() == "pdf":
            return DocumentGenerator.generate_pdf(full_content, title)
        else:
            return DocumentGenerator.generate_docx(full_content, title)
