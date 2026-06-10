# Document Generation Feature

## Overview

The document generation feature allows users to convert chat responses into professional Word documents (DOCX) and PDF files. This is useful for:

- Creating formal reports from technical discussions
- Documenting TML debugging sessions
- Generating specification summaries
- Creating test reports

## How to Use

### Method 1: Automatic Detection

Simply ask for a document in your chat message:

**Examples:**
- "Create a document about TML testing"
- "Generate a PDF report on ADuCM410"
- "Convert this to a doc"
- "Make a report about the test results"
- "Export this as PDF"

When the system detects a document request, download buttons will automatically appear below the response.

### Method 2: Sidebar Export

The sidebar always has a "Document Export" section that allows you to export the last chat response:

1. Get a response from the assistant
2. Go to the sidebar
3. Click "📥 DOCX" or "📄 PDF" to download

Both buttons are ready immediately - no need to click generate first!

### Method 3: Manual Request

You can also explicitly ask:
- "Give me a downloadable document"
- "I need this as a PDF"
- "Create a Word document for this"

## Supported Features

### Content Formatting

The document generator supports markdown formatting:

**Text Formatting:**
- Paragraphs and line breaks
- Bold and italic text (via markdown)
- Headings (H1, H2, H3)

**Lists:**
```markdown
- Bullet point 1
- Bullet point 2

1. Numbered item 1
2. Numbered item 2
```

**Code Blocks:**
```markdown
```tml
testflow example {
    execute = {
        run_test("test1")
    }
}
```
```

**Callouts/Blockquotes:**
```markdown
> Note: This is an important note

> Warning: Critical configuration setting

> Tip: Use this for better performance
```

**Tables:**
Tables are automatically formatted with headers and zebra striping.

### Document Structure

Generated documents include:
- Professional title page with metadata
- Table of contents (for documents with 3+ sections)
- Headers and footers with page numbers
- Consistent styling (Calibri font, blue color scheme)
- Analog Devices branding

### Metadata

Each document includes:
- **Title**: Extracted from content or query
- **Author**: "ADI ChipAgent"
- **Date**: Current date
- **Confidentiality**: "Internal Use Only"

## Technical Details

### Architecture

The document generation system consists of:

1. **DocumentGenerator** (`document_generator.py`): Main service class
   - `detect_document_request()`: Detects user intent
   - `parse_content_to_sections()`: Parses markdown to structured format
   - `generate_docx()`: Creates Word document
   - `generate_pdf()`: Creates PDF document
   - `create_from_llm_response()`: Converts chat response to document

2. **docx_builder** (`docx_builder.py`): Low-level document creation
   - Professional styling and formatting
   - Support for multiple content block types
   - Custom headers, footers, and page layout

### Detection Keywords

The system detects document requests using pattern matching:

**DOCX Keywords:**
- "create a document"
- "generate doc"
- "make a report"
- "convert to document"
- "export as doc"

**PDF Keywords:**
- "create pdf"
- "generate pdf"
- "convert to pdf"
- "export as pdf"

### File Format Support

**DOCX (Word Document):**
- Native format with full formatting support
- Can be opened in Microsoft Word, LibreOffice, Google Docs
- Preserves all styling, tables, code blocks, and callouts

**PDF:**
- Converted from DOCX using `docx2pdf` library
- Platform: Windows (uses COM automation)
- Fallback: PyMuPDF text-based PDF for cross-platform

## Examples

### Example 1: Test Report

**User Query:**
```
Create a document about the TML connectivity test results
```

**Assistant Response:**
```markdown
## TML Connectivity Test Results

### Test Configuration
- Device: ADuCM410
- Platform: V93000
- Test Suite: connectivity_v1.2

### Results
All tests passed:
- Pin continuity: PASS
- Power supply: PASS
- Ground check: PASS

### Code Example
```tml
testflow connectivity {
    execute = {
        run_test("pin_check")
    }
}
```

### Conclusion
Device meets connectivity requirements.
```

**Generated Document:**
- Title page with "TML Connectivity Test Results"
- Table of contents
- Formatted sections with proper headings
- Code block with syntax highlighting
- Professional footer with page numbers

### Example 2: Debugging Guide

**User Query:**
```
Generate a PDF about debugging TML timeout errors
```

**Assistant Response:**
Markdown content with sections on:
- Common causes
- Troubleshooting steps
- Example fixes
- Prevention tips

**Result:**
PDF document with all sections properly formatted and ready to share.

## Troubleshooting

### Issue: Download button doesn't appear

**Solution:** Make sure your query includes document keywords like "create document", "generate pdf", etc.

### Issue: PDF generation fails

**Possible causes:**
- Missing `docx2pdf` library (Windows only)
- COM automation not available
- Permission issues

**Solution:** 
```bash
pip install docx2pdf
```

For non-Windows systems, the system will create a basic PDF using PyMuPDF.

### Issue: Formatting is lost

**Check:**
- Are you using proper markdown syntax?
- Are code blocks properly delimited with triple backticks?
- Are headings using `##` syntax?

### Issue: Unicode characters appear as boxes

**Solution:** The generated documents use standard fonts. Stick to ASCII or common Unicode characters for best compatibility.

## API Reference

### DocumentGenerator Class

```python
from page_indexing_rag.document_generator import DocumentGenerator

# Detect document request
is_request, format_type = DocumentGenerator.detect_document_request(
    "create a document about TML"
)
# Returns: (True, "docx")

# Generate DOCX
docx_bytes = DocumentGenerator.generate_docx(
    content="# Title\n\n## Section\nContent here",
    title="My Document"
)

# Generate PDF
pdf_bytes = DocumentGenerator.generate_pdf(
    content="# Title\n\n## Section\nContent here",
    title="My Document"
)

# Create from LLM response
doc_bytes = DocumentGenerator.create_from_llm_response(
    user_query="Tell me about TML",
    llm_response="## TML Overview\n...",
    format="docx"  # or "pdf"
)
```

### DocumentSpec Structure

```python
from page_indexing_rag.docx_builder import (
    DocumentSpec,
    SectionSpec,
    ParagraphBlock,
    BulletBlock,
    CodeBlock
)

spec = DocumentSpec(
    title="My Report",
    author="ADI ChipAgent",
    date="April 20, 2026",
    sections=[
        SectionSpec(
            heading="Introduction",
            level=2,
            blocks=[
                ParagraphBlock(text="This is a paragraph."),
                BulletBlock(items=["Item 1", "Item 2"]),
                CodeBlock(code="print('hello')", language="python")
            ]
        )
    ]
)
```

## Dependencies

- `python-docx`: Word document creation
- `PyMuPDF` (fitz): PDF manipulation
- `docx2pdf`: DOCX to PDF conversion (Windows)

## Limitations

1. **PDF Conversion**: Works best on Windows. Other platforms use basic PDF generation.
2. **File Size**: Large documents (>50 pages) may take longer to generate.
3. **Images**: Image embedding from markdown is not yet supported.
4. **Tables**: Markdown tables are converted to bullet points (not table format yet).
5. **Styling**: Limited to predefined professional style (no custom fonts/colors).

## Future Enhancements

- [ ] Support for markdown tables
- [ ] Image embedding from URLs or attachments
- [ ] Custom styling options
- [ ] Multi-language support
- [ ] LaTeX equation rendering
- [ ] Chart and graph generation
- [ ] Template selection
- [ ] Batch document generation

## Support

For issues or questions about document generation:
1. Check the test suite: `tests/test_document_generator.py`
2. Run the demo: `python demo_document_generation.py`
3. Review examples in `demo_output/` directory
