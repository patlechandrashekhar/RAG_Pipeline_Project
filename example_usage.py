"""
example_usage.py
Demonstrates all docx_builder capabilities with a sample ChipAgent report.

Run:
    python example_usage.py

Output: example_output/ChipAgent_Sample_Report.docx
"""

from __future__ import annotations

import os
import sys
from datetime import date

# Ensure the src package is importable when running from the project root
_SRC = os.path.join(os.path.dirname(__file__), "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from page_indexing_rag.docx_builder import (
    BulletBlock,
    CalloutBlock,
    CodeBlock,
    DocumentSpec,
    ImageBlock,
    NumberedBlock,
    ParagraphBlock,
    SectionSpec,
    TableSpec,
    build_docx,
)


def main() -> None:
    spec = DocumentSpec(
        title="ADI ChipAgent — Test Engineering Report",
        subtitle="SPI Master Validation on ADuCM410",
        author="Test Engineering Team",
        date=date.today().strftime("%B %d, %Y"),
        confidentiality_label="ADI Internal Use Only",
        page_size="letter",
        include_title_page=True,
        include_toc=True,
        header_text="ADI ChipAgent Report — ADuCM410 SPI Validation",
        sections=[

            # ── 1. Executive Summary ─────────────────────────────────────
            SectionSpec(
                heading="Executive Summary",
                level=1,
                blocks=[
                    ParagraphBlock(
                        "This report summarises the SPI master validation results for the "
                        "ADuCM410 microcontroller (ARM Cortex-M33, 160 MHz via PLL). "
                        "Three test modes were exercised: SPI0 loopback at 8 MHz, "
                        "SPI1 interrupt-driven TX at 4 MHz, and SPI1 polling-based RX at 1 MHz."
                    ),
                    CalloutBlock(
                        text=(
                            "All three test modes passed on revision A silicon. "
                            "Marginal timing observed at SPI0 8 MHz under worst-case PVT — "
                            "see Section 3 for details."
                        ),
                        kind="note",
                        title="RESULT SUMMARY",
                    ),
                ],
            ),

            # ── 2. Device Configuration ──────────────────────────────────
            SectionSpec(
                heading="Device Configuration",
                level=1,
                blocks=[
                    ParagraphBlock(
                        "The test vehicle is the ADuCM410 operating at 160 MHz HCLK "
                        "(PCLK0 = 20 MHz, PCLK1 = 160 MHz). J-Link SWD was used for "
                        "flash programming and debug."
                    ),
                    SectionSpec(
                        heading="Clock Setup",
                        level=2,
                        blocks=[
                            TableSpec(
                                headers=["Parameter", "Value", "Notes"],
                                rows=[
                                    ["HFOSC", "12 MHz", "Internal oscillator"],
                                    ["PLL multiplier", "×160/12", "160 MHz HCLK"],
                                    ["PCLK0", "20 MHz", "÷8 from HCLK"],
                                    ["PCLK1", "160 MHz", "Direct from HCLK"],
                                ],
                                column_widths=[5.5, 4.0, 7.0],
                                zebra=True,
                                header_fill="1F497D",
                                caption="Table 1 — ADuCM410 clock configuration",
                            ),
                        ],
                    ),
                    SectionSpec(
                        heading="SPI Pin Assignments",
                        level=2,
                        blocks=[
                            TableSpec(
                                headers=["Interface", "SCLK", "MISO", "MOSI", "CS"],
                                rows=[
                                    ["SPI0", "P0.0", "P0.1", "P0.2", "P0.3"],
                                    ["SPI1", "P1.4", "P1.5", "P1.6", "P1.7"],
                                ],
                                zebra=True,
                                header_fill="2E75B6",
                                caption="Table 2 — SPI pin mapping",
                            ),
                        ],
                    ),
                ],
            ),

            # ── 3. Test Results ──────────────────────────────────────────
            SectionSpec(
                heading="Test Results",
                level=1,
                blocks=[
                    ParagraphBlock(
                        "The following sub-sections detail each test function defined in "
                        "main.c and the pass/fail outcome observed during characterisation."
                    ),
                    SectionSpec(
                        heading="Spi0LoopbackTest — SPI0 at 8 MHz",
                        level=2,
                        blocks=[
                            ParagraphBlock(
                                "Internal loopback, interrupt-driven RX. "
                                "CPHA=1, CPOL=0 (Mode 2). Baud divider = 9 "
                                "(160 MHz / (2×10) = 8 MHz)."
                            ),
                            BulletBlock(items=[
                                "Transmitted payload: 0xA5, 0x5A, 0xFF, 0x00",
                                "Received payload matched transmitted data on all 1000 iterations",
                                "ISR latency measured: 125–180 ns",
                                "No FIFO overrun or underrun observed",
                            ]),
                            CalloutBlock(
                                text=(
                                    "SCLK eye diagram shows 15% jitter at 3.3 V, 125 °C. "
                                    "Recommend increasing drive strength to 8 mA if operating "
                                    "above 100 °C with trace length >10 cm."
                                ),
                                kind="warning",
                                title="MARGINAL TIMING",
                            ),
                        ],
                    ),
                    SectionSpec(
                        heading="Spi1TxTest — SPI1 at 4 MHz",
                        level=2,
                        blocks=[
                            ParagraphBlock(
                                "Interrupt-driven TX of the string 'ADuCM410' (8 bytes). "
                                "CPHA=1, CPOL=1 (Mode 3). Baud divider = 19."
                            ),
                            NumberedBlock(items=[
                                "Load 'ADuCM410' into SpiTxBuf[8]",
                                "Configure gSpi1Setup with txTrigTransfer=1, continuousTx=1",
                                "Call SpiSetup() and SpiBaud(SPI1, 19)",
                                "Enable TX FIFO interrupt and assert CS",
                                "Verify all 8 bytes shifted out via logic analyser",
                            ]),
                        ],
                    ),
                    SectionSpec(
                        heading="Spi1RxTest — SPI1 at 1 MHz (Active)",
                        level=2,
                        blocks=[
                            ParagraphBlock(
                                "Polling-based RX of 3 bytes. Currently the active test in main(). "
                                "txTrigTransfer is overwritten to 0 at runtime before SpiSetup()."
                            ),
                        ],
                    ),
                ],
            ),

            # ── 4. Register Walkthrough ───────────────────────────────────
            SectionSpec(
                heading="Key Register Configuration",
                level=1,
                blocks=[
                    ParagraphBlock(
                        "The baud rate divider formula and example values used in this project:"
                    ),
                    CodeBlock(
                        code=(
                            "// SpiBaud formula\n"
                            "// SPI clock = PCLK / (2 * (1 + div))\n"
                            "// PCLK1 = 160 MHz\n"
                            "\n"
                            "SpiBaud(SPI0, 9);   // div=9  → 160/(2×10) = 8 MHz\n"
                            "SpiBaud(SPI1, 19);  // div=19 → 160/(2×20) = 4 MHz\n"
                            "SpiBaud(SPI1, 79);  // div=79 → 160/(2×80) = 1 MHz\n"
                        ),
                        language="C",
                        caption="SPI baud rate configuration",
                    ),
                    ParagraphBlock(
                        "The SPI0 setup struct initialised in ADuCM410_Setup.c:"
                    ),
                    CodeBlock(
                        code=(
                            "SpiSetup_t gSpi0Setup = {\n"
                            "    .master           = 1,\n"
                            "    .cpol             = 0,  // CPHA=1, CPOL=0 → Mode 2\n"
                            "    .cpha             = 1,\n"
                            "    .txTrigTransfer   = 1,\n"
                            "    .continuousTx     = 1,\n"
                            "    .irqMode          = IRQMODE_TX1RX1,\n"
                            "};\n"
                        ),
                        language="C",
                        caption="gSpi0Setup struct — SPI0 loopback configuration",
                    ),
                ],
            ),

            # ── 5. Recommendations ───────────────────────────────────────
            SectionSpec(
                heading="Recommendations & Next Steps",
                level=1,
                blocks=[
                    ParagraphBlock(
                        "Based on the validation results, the following actions are recommended "
                        "before tape-out sign-off:"
                    ),
                    NumberedBlock(items=[
                        "Increase SPI0 MOSI/SCLK GPIO drive strength to 8 mA for high-temperature operation",
                        "Add FIFO depth monitoring in Spi1RxTest to detect potential buffer stalls",
                        "Characterise timing margins across the full PVT corner matrix (SS/FF/TT × −40/27/125°C)",
                        "Enable DMA-driven SPI for transfers >32 bytes to reduce CPU interrupt overhead",
                        "Validate ADIN6310 TSN switch SPI register access using Spi1TxTest framework",
                    ]),
                    CalloutBlock(
                        text="Spi1RxTest is the only test active in main() — re-enable Spi0LoopbackTest and Spi1TxTest before regression suite execution.",
                        kind="important",
                    ),
                ],
            ),

        ],
    )

    # Flatten nested SectionSpec objects to top-level (builder handles nesting via levels)
    flat_sections: list = []
    for sec in spec.sections:
        flat_sections.append(sec)
        for block in list(sec.blocks):
            if isinstance(block, SectionSpec):
                sec.blocks.remove(block)
                flat_sections.append(block)

    spec.sections = flat_sections

    out_dir = os.path.join(os.path.dirname(__file__), "example_output")
    out_path = os.path.join(out_dir, "ChipAgent_Sample_Report.docx")

    result = build_docx(spec, out_path)
    print(f"Document created: {result}")


if __name__ == "__main__":
    main()
