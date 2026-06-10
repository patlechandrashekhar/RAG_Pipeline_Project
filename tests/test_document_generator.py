"""Test document generation functionality."""

import sys
from pathlib import Path

import pytest

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from page_indexing_rag.document_generator import DocumentGenerator


class TestDocumentGenerator:
    """Test document generation service."""

    def test_detect_document_request_docx(self):
        """Test detection of DOCX generation requests."""
        test_cases = [
            ("create a document about TML testing", True, "docx"),
            ("generate a doc with test results", True, "docx"),
            ("convert this to a document", True, "docx"),
            ("make a report about the findings", True, "docx"),
        ]

        for message, expected_is_request, expected_format in test_cases:
            is_request, format_type = DocumentGenerator.detect_document_request(message)
            assert is_request == expected_is_request, f"Failed for: {message}"
            if expected_format:
                assert format_type == expected_format, f"Wrong format for: {message}"

    def test_detect_document_request_pdf(self):
        """Test detection of PDF generation requests."""
        test_cases = [
            ("create a pdf about TML testing", True, "pdf"),
            ("generate pdf with test results", True, "pdf"),
            ("convert this to pdf", True, "pdf"),
            ("export as pdf", True, "pdf"),
        ]

        for message, expected_is_request, expected_format in test_cases:
            is_request, format_type = DocumentGenerator.detect_document_request(message)
            assert is_request == expected_is_request, f"Failed for: {message}"
            assert format_type == expected_format, f"Wrong format for: {message}"

    def test_detect_no_document_request(self):
        """Test that regular queries are not detected as document requests."""
        test_cases = [
            "What is TML?",
            "How does the V93000 work?",
            "Explain ADuCM410 pinout",
            "Debug this test program",
        ]

        for message in test_cases:
            is_request, _ = DocumentGenerator.detect_document_request(message)
            assert not is_request, f"False positive for: {message}"

    def test_parse_content_to_sections_simple(self):
        """Test parsing simple markdown content."""
        content = """# Test Report

## Introduction
This is a test report about TML debugging.

## Findings
- Finding 1
- Finding 2
- Finding 3

## Conclusion
The test was successful.
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        assert doc_spec.title == "Test Report"
        assert doc_spec.author == "ADI ChipAgent"
        assert len(doc_spec.sections) > 0
        assert doc_spec.include_title_page is True

    def test_parse_content_with_code_blocks(self):
        """Test parsing content with code blocks."""
        content = """## TML Example

Here is a TML code example:

```tml
testflow test_basic {
    setup = {
        Vdd = 3.3V
    }
    execute = {
        run_test("basic_connectivity")
    }
}
```

This code demonstrates a basic test flow.
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        assert len(doc_spec.sections) > 0
        # Check that code block was parsed
        section = doc_spec.sections[0]
        has_code_block = any(
            hasattr(block, "code") for block in section.blocks
        )
        assert has_code_block, "Code block not found in parsed sections"

    def test_parse_content_with_lists(self):
        """Test parsing content with bullet and numbered lists."""
        content = """## Test Results

### Bullet List
- Item 1
- Item 2
- Item 3

### Numbered List
1. First step
2. Second step
3. Third step
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        assert len(doc_spec.sections) >= 2
        # Verify lists were parsed
        all_blocks = []
        for section in doc_spec.sections:
            all_blocks.extend(section.blocks)

        from page_indexing_rag.docx_builder import BulletBlock, NumberedBlock

        has_bullet = any(isinstance(block, BulletBlock) for block in all_blocks)
        has_numbered = any(isinstance(block, NumberedBlock) for block in all_blocks)

        assert has_bullet, "Bullet list not found"
        assert has_numbered, "Numbered list not found"

    def test_parse_content_with_alt_list_markers(self):
        """Test parsing content with alternate bullet and numbering styles."""
        content = """## Implementation Plan

### Bullet Markers
• Item A
• Item B

### Numbering Markers
1) First action
2) Second action
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        all_blocks = []
        for section in doc_spec.sections:
            all_blocks.extend(section.blocks)

        from page_indexing_rag.docx_builder import BulletBlock, NumberedBlock

        has_bullet = any(isinstance(block, BulletBlock) for block in all_blocks)
        has_numbered = any(isinstance(block, NumberedBlock) for block in all_blocks)

        assert has_bullet, "Bullet list with unicode marker not found"
        assert has_numbered, "Numbered list with ')' marker not found"

    def test_generate_docx(self):
        """Test DOCX generation."""
        content = """# Test Document

## Section 1
This is a test section with some content.

## Section 2
Another section with more information.
"""
        docx_bytes = DocumentGenerator.generate_docx(content, "Test Title")

        assert docx_bytes is not None
        assert len(docx_bytes) > 0
        assert isinstance(docx_bytes, bytes)

    def test_create_from_llm_response_docx(self):
        """Test creating DOCX from LLM response."""
        user_query = "Tell me about TML testing"
        llm_response = """## TML Testing Overview

TML (Test Markup Language) is used for writing test programs on V93000 ATE.

### Key Features
- Declarative syntax
- Hardware abstraction
- Reusable test modules

### Example
```tml
testflow basic {
    execute = {
        run_test("connectivity")
    }
}
```
"""
        docx_bytes = DocumentGenerator.create_from_llm_response(
            user_query, llm_response, format="docx"
        )

        assert docx_bytes is not None
        assert len(docx_bytes) > 0
        assert isinstance(docx_bytes, bytes)

    def test_parse_callout_blocks(self):
        """Test parsing callout/blockquote content."""
        content = """## Important Notes

> Warning: This is a critical configuration setting.

> Tip: Use this approach for better performance.
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        all_blocks = []
        for section in doc_spec.sections:
            all_blocks.extend(section.blocks)

        from page_indexing_rag.docx_builder import CalloutBlock

        callouts = [block for block in all_blocks if isinstance(block, CalloutBlock)]
        assert len(callouts) > 0, "Callout blocks not found"

    def test_empty_content(self):
        """Test handling of empty content."""
        content = ""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        assert doc_spec.title == "Generated Document"
        # Should still have basic structure
        assert doc_spec.author == "ADI ChipAgent"

    def test_title_extraction(self):
        """Test automatic title extraction from content."""
        content = """# My Custom Title

Some content here.
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(content)

        assert doc_spec.title == "My Custom Title"

    def test_title_override(self):
        """Test explicit title override."""
        content = """# Original Title

Content.
"""
        doc_spec = DocumentGenerator.parse_content_to_sections(
            content, title="Override Title"
        )

        assert doc_spec.title == "Override Title"
