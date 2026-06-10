"""
Demo script for document generation functionality.

This script demonstrates how to use the DocumentGenerator to create
professional Word documents and PDFs from content.
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from page_indexing_rag.document_generator import DocumentGenerator


def demo_basic_document():
    """Demo: Generate a basic document."""
    print("=" * 60)
    print("Demo 1: Basic Document Generation")
    print("=" * 60)

    content = """# TML Testing Report

## Executive Summary
This report provides an overview of TML (Test Markup Language) testing
performed on the V93000 ATE platform for ADI devices.

## Test Setup
The following test configuration was used:

### Equipment
- Advantest V93000 ATE
- AVI64 instrument cards
- DPVS power supply cards
- HVIN high-voltage cards

### Device Under Test
- Device: ADuCM410
- Package: 64-pin LFCSP
- Test board: Rev 2.1

## Test Results

### Connectivity Tests
All connectivity tests passed successfully:
- Pin continuity: PASS
- Power supply: PASS
- Ground connections: PASS

### Functional Tests
1. Digital I/O test: PASS
2. Analog input test: PASS
3. Communication test (SPI): PASS
4. Power consumption test: PASS

## TML Code Example

```tml
testflow connectivity_check {
    setup = {
        Vdd = 3.3V
        Vss = 0V
    }
    execute = {
        run_test("pin_continuity")
        run_test("power_supply_check")
        run_test("ground_check")
    }
    limits = {
        continuity_resistance < 10 ohms
    }
}
```

## Issues and Recommendations

> Warning: Pin 23 showed slightly elevated resistance (8.5 ohms).
> Monitor this in production testing.

> Tip: Consider adding temperature monitoring for long-duration tests.

## Conclusion
All tests completed successfully. Device meets specifications and is
ready for production deployment.

## Next Steps
1. Validate on 10 additional units
2. Update test limits database
3. Generate characterization report
"""

    # Test detection
    test_messages = [
        "create a document about this test",
        "generate a pdf report",
        "convert this to doc",
    ]

    print("\nTesting document request detection:")
    for msg in test_messages:
        is_request, format_type = DocumentGenerator.detect_document_request(msg)
        print(f"  '{msg}' -> Request: {is_request}, Format: {format_type}")

    # Generate DOCX
    print("\nGenerating DOCX...")
    docx_bytes = DocumentGenerator.generate_docx(content, "TML Testing Report")
    print(f"[OK] DOCX generated: {len(docx_bytes):,} bytes")

    # Save to file
    output_path = Path("demo_output")
    output_path.mkdir(exist_ok=True)

    docx_file = output_path / "tml_testing_report.docx"
    with open(docx_file, "wb") as f:
        f.write(docx_bytes)
    print(f"[OK] Saved to: {docx_file}")

    # Generate PDF
    print("\nGenerating PDF...")
    try:
        pdf_bytes = DocumentGenerator.generate_pdf(content, "TML Testing Report")
        print(f"[OK] PDF generated: {len(pdf_bytes):,} bytes")

        pdf_file = output_path / "tml_testing_report.pdf"
        with open(pdf_file, "wb") as f:
            f.write(pdf_bytes)
        print(f"[OK] Saved to: {pdf_file}")
    except Exception as e:
        print(f"[ERROR] PDF generation error: {e}")
        print("   (This is normal on some systems without proper PDF converters)")


def demo_llm_response():
    """Demo: Generate document from simulated LLM response."""
    print("\n" + "=" * 60)
    print("Demo 2: Document from LLM Response")
    print("=" * 60)

    user_query = "Explain how to debug TML test programs"
    llm_response = """## TML Debugging Guide

### Common Issues and Solutions

#### 1. Syntax Errors
TML syntax errors are often caught at compile time. Check for:
- Missing semicolons
- Unmatched braces
- Typos in keywords

```tml
// Incorrect
testflow bad_example {
    execute =
        run_test("test1")  // Missing braces
    }
}

// Correct
testflow good_example {
    execute = {
        run_test("test1");
    }
}
```

#### 2. Runtime Failures
For runtime failures:
1. Check instrument connections
2. Verify voltage levels
3. Review timing constraints
4. Examine test limits

#### 3. Debugging Tools
- Use `print()` statements for variable inspection
- Enable verbose logging in test suite
- Review waveforms in data log
- Check instrument status registers

> Important: Always verify hardware connections before debugging software.

### Best Practices
- Use descriptive test names
- Add comments for complex logic
- Implement error handling
- Log intermediate results
"""

    print(f"User Query: {user_query}")
    print(f"LLM Response: {len(llm_response)} characters")

    # Generate document
    print("\nGenerating DOCX from LLM response...")
    docx_bytes = DocumentGenerator.create_from_llm_response(
        user_query, llm_response, format="docx"
    )
    print(f"[OK] DOCX generated: {len(docx_bytes):,} bytes")

    output_path = Path("demo_output")
    output_path.mkdir(exist_ok=True)

    docx_file = output_path / "tml_debugging_guide.docx"
    with open(docx_file, "wb") as f:
        f.write(docx_bytes)
    print(f"[OK] Saved to: {docx_file}")


def demo_content_parsing():
    """Demo: Show content parsing capabilities."""
    print("\n" + "=" * 60)
    print("Demo 3: Content Parsing")
    print("=" * 60)

    content = """# Device Specifications

## Electrical Characteristics

### Operating Conditions
- Supply Voltage (Vdd): 2.7V to 3.6V
- Operating Temperature: -40°C to +85°C
- Storage Temperature: -55°C to +150°C

### Digital I/O
- High-level Input (Vih): 0.7 × Vdd
- Low-level Input (Vil): 0.3 × Vdd
- Output Drive Strength: 4mA typical

## Pin Configuration
1. Pin 1: Vdd (Power Supply)
2. Pin 2: GND (Ground)
3. Pin 3: SCLK (SPI Clock)
4. Pin 4: MOSI (SPI Master Out)
5. Pin 5: MISO (SPI Master In)

## Communication Protocols

### SPI Interface
The device supports SPI Mode 0 and Mode 3:
- Mode 0: CPOL = 0, CPHA = 0
- Mode 3: CPOL = 1, CPHA = 1

Clock frequency up to 10 MHz.

### I2C Interface (Secondary)
- Standard mode: 100 kHz
- Fast mode: 400 kHz
- 7-bit addressing only

> Note: I2C pull-up resistors are not integrated. External 4.7kΩ resistors recommended.
"""

    print("Parsing markdown content...")
    doc_spec = DocumentGenerator.parse_content_to_sections(content)

    print(f"[OK] Title: {doc_spec.title}")
    print(f"[OK] Author: {doc_spec.author}")
    print(f"[OK] Sections: {len(doc_spec.sections)}")

    for i, section in enumerate(doc_spec.sections, 1):
        print(f"  Section {i}: {section.heading} (Level {section.level})")
        print(f"    Blocks: {len(section.blocks)}")


if __name__ == "__main__":
    print("\nDocument Generation Demo\n")

    try:
        demo_basic_document()
        demo_llm_response()
        demo_content_parsing()

        print("\n" + "=" * 60)
        print("All demos completed successfully!")
        print("=" * 60)
        print("\nGenerated documents are in the 'demo_output' directory.")
        print("Open them in Microsoft Word or compatible viewer.\n")

    except Exception as e:
        print(f"\nError: {e}")
        import traceback

        traceback.print_exc()
