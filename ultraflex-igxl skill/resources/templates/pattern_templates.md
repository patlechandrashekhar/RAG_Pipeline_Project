# UltraFlex ATP Pattern Templates

Generic ATP pattern templates for the UltraFlex/IG-XL platform.
UltraFlex patterns target UltraPin800/UltraPin1600/UltraPin4000
digital boards with features like SBC drive format, dual-2X timing,
256M vector memory, and 255-cycle pipeline.

---

## Table of Contents

1. [ATP File Structure (UltraFlex)](#atp-file-structure-ultraflex)
2. [Vector Notation](#vector-notation)
3. [Pattern Opcodes (UltraFlex-specific)](#pattern-opcodes-ultraflex-specific)
4. [PVS File](#pvs-file)
5. [SPI Write Pattern (bit-bang)](#1-spi-write-pattern-bit-bang)
6. [SPI Read Pattern (with capture)](#2-spi-read-pattern-with-capture)
7. [SPI with SBC Clock (UltraFlex-specific)](#3-spi-with-sbc-clock-ultraflex-specific)
8. [GPIO Check Pattern](#4-gpio-check-pattern)
9. [JTAG / ARM SWD Pattern](#5-jtag--arm-swd-pattern)
10. [Keep-Alive Clock Pattern](#6-keep-alive-clock-pattern)
11. [Frequency Counter Pattern](#7-frequency-counter-pattern)
12. [Scan Chain with DSSC (UltraFlex-specific)](#8-scan-chain-with-dssc-ultraflex-specific)
13. [Flat SPI Pattern System](#9-flat-spi-pattern-system)
14. [Pattern with Pipeline Flush (UltraFlex-specific)](#10-pattern-with-pipeline-flush-ultraflex-specific)
15. [DUT Reset Pattern](#11-dut-reset-pattern)

---

## ATP File Structure (UltraFlex)

```
opcode_mode = single;        // or dual, dual_2x
digital_inst = hsd;          // target UltraPin800
check_vm_min_size = 1000;    // optional: minimum vector memory check

import tset ts_dut;
import subr my_subroutine;   // optional

vector ($tset, SCLK, DIN, CSB, DOUT)
{
start_label MyPattern:
> ts_dut  0 0 1 X ;
halt;
}
```

### Key Header Differences from MicroFlex

- `digital_inst = hsd` (UltraFlex) vs. none (MicroFlex).
- `check_vm_min_size` for vector memory validation.
- Support for `opcode_mode = dual` and `dual_2x`.

---

## Vector Notation

Same as MicroFlex plus:

| Symbol | Meaning                                     |
|--------|---------------------------------------------|
| `0`    | Drive logic low                             |
| `1`    | Drive logic high                            |
| `X`    | Don't care (no compare)                     |
| `H`    | Expect logic high                           |
| `L`    | Expect logic low                            |
| `C`    | Capture (store result)                      |
| `Z`    | Expect high-impedance                       |
| `M`    | Midband (three-level driver, UltraPin800)   |

### Drive Formats (UltraPin800)

- `NR` - Non-Return
- `RH` - Return-High
- `RL` - Return-Low
- `SBC` - Surround By Complement (commonly used for clock pins)
- `SBH` - Surround By High
- `SBL` - Surround By Low
- `STAY` - Maintain previous state
- 2X variants for dual-2X mode: `NR_2X`, `RH_2X`, etc.

---

## Pattern Opcodes (UltraFlex-specific)

```
// Flow control
halt                    // stop
repeat <n>              // repeat next vector N times
loop <label>, <n>       // loop to label N times
end_loop                // end loop
jump <label>            // unconditional jump
call <subr>             // call subroutine
return                  // return from subroutine
match <label>, <n>      // wait for compare match

// Counter stack (UltraFlex-specific)
set_msb <value>         // set counter MSB
set <value>             // set counter value
push_loop               // push loop counter
pop_loop                // pop and decrement

// DSSC control
stv                     // store this vector (capture marker)
mask                    // mask comparison

// CPU handshake
cpuA, cpuB, cpuC, cpuD  // flags for VBT-pattern sync

// Pipeline (UltraFlex-specific)
pipe_minus              // pipeline flush (255 cycles on UltraPin800)
```

---

## PVS File

Same as MicroFlex. A companion `.pvs` file is required for every
`.pat` file. The `.pvs` defines pin groups, time set associations,
and pattern list entries that reference the ATP pattern.

---

## 1. SPI Write Pattern (bit-bang)

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_spi;

vector ($tset, SCLK, MOSI, CSN, MISO)
{
start_label SPI_Write:
> ts_spi  0 0 1 X ;     // idle: CSN high
> ts_spi  0 0 0 X ;     // CSN low - start

// Bit[N-1] MSB
> ts_spi  0 <bit> 0 X ; // setup data
> ts_spi  1 <bit> 0 X ; // SCLK rising edge

// ... repeat for each bit ...

// Bit[0] LSB
> ts_spi  0 <bit> 0 X ;
> ts_spi  1 <bit> 0 X ;

> ts_spi  0 0 1 X ;     // CSN high - end
halt;
}
```

---

## 2. SPI Read Pattern (with capture)

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_spi;

vector ($tset, SCLK, MOSI, CSN, MISO)
{
start_label SPI_Read:
> ts_spi  0 0 1 X ;     // idle

// --- Command/address phase ---
> ts_spi  0 0 0 X ;     // CSN low
> ts_spi  0 1 0 X ;     // R/W = 1 (read)
> ts_spi  1 1 0 X ;
// <address bits>

// --- Data capture phase ---
> ts_spi  0 0 0 X ;     // setup
stv > ts_spi  1 0 0 C ; // capture MISO on rising edge
// ... repeat for each data bit ...

> ts_spi  0 0 1 X ;     // CSN high
halt;
}
```

---

## 3. SPI with SBC Clock (UltraFlex-specific)

Uses SBC (Surround By Complement) drive format for cleaner clock
edges. In the time set definition, SCLK uses SBC format with edges
at 25% and 75% of the period.

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_spi_sbc;

// In time set definition, SCLK uses SBC format:
// Drive format: SBC with edges at 25% and 75% of period

vector ($tset, SCLK, MOSI, CSN, MISO)
{
start_label SPI_Write_SBC:
> ts_spi_sbc  0 0 1 X ;   // idle
> ts_spi_sbc  0 0 0 X ;   // CSN low
// With SBC format, SCLK toggles automatically around data:
// drive=0 -> low-high-low within cycle
// drive=1 -> high-low-high within cycle
> ts_spi_sbc  1 <bit> 0 X ;  // SCLK pulse with data
> ts_spi_sbc  1 <bit> 0 X ;
// ... repeat ...
> ts_spi_sbc  0 0 1 X ;   // CSN high
halt;
}
```

---

## 4. GPIO Check Pattern

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_gpio;

vector ($tset, GPIO_0, GPIO_1, GPIO_2, GPIO_3)
{
start_label GPIO_AllHigh:
> ts_gpio  H H H H ;    // expect all high
> ts_gpio  H H H H ;
> ts_gpio  H H H H ;
halt;

GPIO_AllLow:
> ts_gpio  L L L L ;    // expect all low
> ts_gpio  L L L L ;
> ts_gpio  L L L L ;
halt;
}
```

---

## 5. JTAG / ARM SWD Pattern

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_jtag;

vector ($tset, SWCLK, SWDIO)
{
start_label SWD_Init:
// Line reset: 50+ clocks with SWDIO high
repeat 56
> ts_jtag  0 1 ;
> ts_jtag  1 1 ;
end_repeat

// JTAG-to-SWD switch sequence (0xE79E)
> ts_jtag  0 0 ;    // bit 0
> ts_jtag  1 0 ;
> ts_jtag  0 1 ;    // bit 1
> ts_jtag  1 1 ;
// ... remaining 14 bits of 0xE79E ...

// Another line reset
repeat 56
> ts_jtag  0 1 ;
> ts_jtag  1 1 ;
end_repeat

// Idle
> ts_jtag  0 0 ;
> ts_jtag  1 0 ;
halt;
}
```

---

## 6. Keep-Alive Clock Pattern

Same concept as MicroFlex but targeting UltraPin800:

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_clk;

vector ($tset, DUT_CLK)
{
start_label keepalive:
> ts_clk  0 ;
> ts_clk  1 ;
> ts_clk  0 ;
> ts_clk  1 ;
halt;
}
```

---

## 7. Frequency Counter Pattern

UltraFlex UltraPin800 has a built-in frequency counter per channel:

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_freq;

vector ($tset, CLK_OUT)
{
start_label FreqCount:
// Run enough cycles for frequency measurement.
// Counter accumulates edges during pattern execution.
repeat 10000
> ts_freq  X ;    // don't care - just count edges
end_repeat
halt;
}
```

VBT usage:

```vba
TheHdw.Digital.Patterns.Pat("freq_count.pat").Run ""
Dim freq As New SiteDouble
freq = TheHdw.Pins("CLK_OUT").FrequencyMeasurement.Read
```

---

## 8. Scan Chain with DSSC (UltraFlex-specific)

UltraFlex scan uses DSSC with external text files for scan vectors:

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_scan;

vector ($tset, SCAN_CLK, SCAN_EN, SCAN_IN, SCAN_OUT)
{
start_label DSSC_Scan:
// Enable scan mode
> ts_scan  0 1 0 X ;

// Shift in/out using DSSC source/capture.
// SCAN_IN data loaded from DSPWave via DSSC source.
// SCAN_OUT captured to DSPWave via DSSC capture.
repeat <num_scan_bits>
> ts_scan  0 1 0 X ;         // setup phase
stv > ts_scan  1 1 0 C ;     // clock + capture
end_repeat

// Disable scan mode
> ts_scan  0 0 0 X ;
halt;
}
```

VBT setup for DSSC scan:

```vba
' Load scan data from text files
Dim scanSrc As New DSPWave
Dim scanCap As New DSPWave
Dim scanMask As New DSPWave
scanSrc.LoadFromFile ".\patterns\ScanIN.txt"

' Configure DSSC source
With TheHdw.DSSC("SCAN_IN").Pattern(patName).Source.Signals
    .Add "scanData"
    .Item("scanData").WaveDefinitionName = "scanData"
    .Item("scanData").SampleSize = numBits
    .Item("scanData").LoadSamples
End With

' Configure DSSC capture
With TheHdw.DSSC.Pins("SCAN_OUT").Pattern(patName).Capture
    .Signals.Add "scanCapture"
    .Signals("scanCapture").SampleSize = numBits
    .Signals("scanCapture").LoadSettings
End With

' Run pattern, retrieve and compare
TheHdw.Digital.Patterns.Pat(patName).Run ""
Set scanCap = TheHdw.DSSC.Pins("SCAN_OUT") _
    .Pattern(patName).Capture.Signals("scanCapture").DSPWave
' Compare: XOR captured with expected, AND with mask, CalcSum
```

---

## 9. Flat SPI Pattern System

Same concept as MicroFlex - pre-compiled labels for fast register
access:

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_spi;

vector ($tset, SCLK, MOSI, CSN, MISO)
{
WriteSPI_16_01_4010:
> ts_spi  0 0 1 X ;
> ts_spi  0 0 0 X ;
// ... 16 bit-bang vectors ...
> ts_spi  0 0 1 X ;
halt;

WriteSPI_16_01_4020:
> ts_spi  0 0 1 X ;
> ts_spi  0 0 0 X ;
// ... 16 bit-bang vectors ...
> ts_spi  0 0 1 X ;
halt;
}
```

---

## 10. Pattern with Pipeline Flush (UltraFlex-specific)

UltraPin800 has a 255-cycle pipeline. Use `pipe_minus` when
immediate effect is needed:

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_dut;

vector ($tset, SCLK, DIN, CSB, DOUT)
{
start_label PipelineExample:
> ts_dut  0 0 1 X ;
// ... pattern vectors ...

// Force pipeline flush before critical timing section
pipe_minus
> ts_dut  0 0 0 X ;    // This vector executes immediately after flush
// ... time-critical vectors ...
halt;
}
```

---

## 11. DUT Reset Pattern

```
opcode_mode = single;
digital_inst = hsd;

import tset ts_dut;

vector ($tset, RESETN, SCLK, MOSI, CSN, MISO)
{
start_label DUT_Reset:
// Assert reset (active low)
> ts_dut  0 0 0 1 X ;
repeat 100
> ts_dut  0 0 0 1 X ;   // hold reset for 100 cycles
end_repeat

// Release reset
> ts_dut  1 0 0 1 X ;
repeat 1000
> ts_dut  1 0 0 1 X ;   // wait for DUT to initialize
end_repeat
halt;
}
```
