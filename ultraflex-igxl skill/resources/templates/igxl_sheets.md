# UltraFlex IG-XL Sheet Format Templates

UltraFlex IG-XL sheet templates use `DT<SheetType>` headers containing
`platform=Jaguar` metadata. These differ from MicroFlex `DFF` headers.
UltraFlex supports additional sheets for concurrent test,
protocol-aware interfaces, and sub-programs.

---

## Table of Contents

1. [Sheet Header Format](#sheet-header-format)
2. [Test Instances Sheet](#1-test-instances-sheet)
3. [Flow Table](#2-flow-table)
4. [Pin Map](#3-pin-map)
5. [Channel Map](#4-channel-map)
6. [Pin Levels](#5-pin-levels)
7. [Time Sets](#6-time-sets)
8. [Global Specs](#7-global-specs)
9. [Job List](#8-job-list)
10. [DC Specs and AC Specs](#9-dc-specs-and-ac-specs)
11. [Load Board File](#10-load-board-file-load)
12. [Sub-Program Structure](#11-sub-program-structure-ultraflex-specific)
13. [Tester Configuration](#12-tester-configuration)

---

## Sheet Header Format

UltraFlex sheets use `DT<SheetType>` with embedded key-value metadata
separated by colons. The general form is:

```
DT<SheetType>,version=<ver>:platform=Jaguar:toprow=<n>:leftcol=<n>:rightcol=<n>
```

Full example:

```
DTTestInstancesSheet,version=2.4:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

Common sheet type prefixes:

| Prefix                   | Purpose                |
|--------------------------|------------------------|
| `DTTestInstancesSheet`   | Test instances         |
| `DTFlowtableSheet`       | Flow table             |
| `DTPinMap`               | Pin map                |
| `DTChanMap`              | Channel map            |
| `DTLevelSheet`           | Pin levels             |
| `DTTimesetBasicSheet`    | Time sets              |
| `DTGlobalSpecSheet`      | Global specs           |
| `DTDCSpecSheet`          | DC specifications      |
| `DTACSpecSheet`          | AC specifications      |
| `DTJobListSheet`         | Job list               |
| `DTHomeSheet`            | Home / summary sheet   |

---

## 1. Test Instances Sheet

**Header:**

```
DTTestInstancesSheet,version=2.4:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Defines every test that the program can execute. Each row
maps a test name to its type (VBT, functional, DC, etc.), spec
selectors, timing, levels, and up to 100 arguments.

**Column headers:**

```
Test Name | Type | Name | Called As | DC Specs Category | DC Specs Selector | AC Specs Category | AC Specs Selector | Time Sets | Edge Sets | Pin Levels | Mixed Signal Timing | Overlay | Arg0 | ... | Arg99 | Comment
```

**Example rows:**

```
TF_InitSetup       VBT  TF_InitSetup               Cat0  vddio_nom  Cat0  Sel0  TSB_gen    Levels_gen
TF_Continuity      VBT  TF_Continuity              Cat0  vddio_nom  Cat0  Sel0  TSB_gen    Levels_gen
TF_SetLevels       VBT  TF_SetLevels               Cat0  vddio_nom  Cat0  Sel0  TSB_gen    Levels_gen   ConditionSetup  1
Trim_LNA_LDO       VBT  Trim_LNA_LDO               Cat0  vddio_nom  Cat0  Sel0  TSB_gen    Levels_gen
```

**Key differences from MicroFlex:**

- DC and AC Specs use Category/Selector pairs (e.g., `Cat0` / `vddio_nom`).
- Arguments can reference spec variables rather than hard-coded values.
- The `DT` header carries version and platform metadata.

---

## 2. Flow Table

**Header:**

```
DTFlowtableSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Controls the execution order of tests, binning, and
branching logic. Same column structure as MicroFlex but with
additional UltraFlex opcodes.

UltraFlex-specific opcodes:

| Opcode            | Description                        |
|-------------------|------------------------------------|
| `concurrent`      | Start a concurrent sub-flow        |
| `concurrent-end`  | End the concurrent block           |
| `call`            | Call a named sub-flow              |

**Column headers:**

```
Label | Enable | Opcode | Parameter | TName | TNum | LoLim | HiLim | Units | Pass | Fail | Pass | Fail | Comment
```

**Example -- calibration phase:**

```
         Calok     set-error-bin                                                                31    99    31    99
         Calibrate Test          TF_InitSetup
                   Use-Limit     TF_InitSetup        SetupResult          100   0       1                 8     8     Fail
         Calok     Test          TF_Continuity
                   Use-Limit     TF_Continuity       DigGP26SupDioV       200   2.9     3.1               8     8     Fail
                   Use-Limit     TF_Continuity       DigGP26GndDioV       201   -1.1    -0.3              8     8     Fail
```

**Example -- production phase with concurrent test:**

```
         Calok     Test          TF_SetLevels1
                   Use-Limit     TF_SetLevels1       Vdd                  300   3.1     3.5               3     3     Fail
         Calok     concurrent    CSF_DigitalTests
         Calok     concurrent    CSF_AnalogTests
         Calok     concurrent-end
         Calok     Test          TF_SupplyCurrent
                   Use-Limit     TF_SupplyCurrent    Idd_Active           400   0       50      uA        4     4     Fail
```

Notes:
- `concurrent` / `concurrent-end` blocks allow multiple sub-flows to
  run simultaneously across different instrument resources.
- Enable words (e.g., `Calok`, `Calibrate`) control which sections
  execute based on the load-board enable settings.

---

## 3. Pin Map

**Header:**

```
DTPinMap,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Maps logical pin names to types and organizes them into
groups. UltraFlex pin maps use the `grp_` prefix for group names.

**Column headers:**

```
Group Name | Pin Name | Type | Comment
```

**Example rows -- individual pins:**

```
                        VDD_ANA         Analog    Analog supply
                        VDD_DIG         Analog    Digital supply
                        VDD_IO          Analog    IO supply
                        HSD_SCLK        I/O       SPI clock
                        HSD_MOSI        I/O       SPI MOSI
                        HSD_MISO        I/O       SPI MISO
                        HSD_CSN         I/O       SPI chip select
                        HSD_GPIO_0      I/O       GPIO port 0 bit 0
                        HSD_RESETN      I/O       Active-low reset
                        RLY_K1_MUX      Utility   Mux relay control
```

**Example rows -- pin groups:**

```
grp_SUPPLY_PINS         VDD_ANA         Analog
grp_SUPPLY_PINS         VDD_DIG         Analog
grp_SUPPLY_PINS         VDD_IO          Analog
grp_DIG_PINS            HSD_SCLK        I/O
grp_DIG_PINS            HSD_MOSI        I/O
grp_DIG_PINS            HSD_MISO        I/O
grp_DIG_PINS            HSD_CSN         I/O
grp_OS_1                HSD_SCLK        I/O       Checkerboard set A (odd)
grp_OS_1                HSD_MISO        I/O
grp_OS_2                HSD_MOSI        I/O       Checkerboard set B (even)
grp_OS_2                HSD_CSN         I/O
grp_ALL_RELAYS          RLY_K1_MUX      Utility
```

**UltraFlex naming conventions:**

| Prefix        | Instrument / Purpose                              |
|---------------|---------------------------------------------------|
| `HSD_`        | UltraPin800 digital pins                          |
| `DC30_`       | DC-30 DCVI pins (add `_HS` suffix for high-sense) |
| `UVI80_`      | UltraVI80 pins                                    |
| `BBAC_`       | Broadband AC pins (BBAC-15)                        |
| `UPAC_`       | UltraPAC pins                                     |
| `NVM_DIFF_`   | Differential meter pins                            |
| `RLY_` / `RL_`| Relay control                                     |
| `grp_`        | Pin groups                                         |

---

## 4. Channel Map

**Header:**

```
DTChanMap,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Maps logical pin names to physical tester channels.
UltraFlex channel notation uses `<slot>.<channel>` or
`<slot>.<resource>.<detail>`.

**Column headers:**

```
Pin Name | Pkg Pin | Type | Site0 | Site1 | Comment
```

**Example rows:**

```
VDD_ANA        A1       DCVI     12.ch0.dgs0        12.ch4.dgs0        UltraVI80
VDD_DIG        A2       DCVI     12.ch1.dgs0        12.ch5.dgs0        UltraVI80
VDD_IO         A3       DCVI     5.sense0.dgs1      5.sense2.dgs3      DC-30
HSD_SCLK       B1       I/O      3.ch0              3.ch64             UltraPin800
HSD_MOSI       B2       I/O      3.ch1              3.ch65             UltraPin800
HSD_MISO       B3       I/O      3.ch2              3.ch66             UltraPin800
BBAC_SRC_P     C1       BBAC     8.ch0              8.ch2              BBAC-15
BBAC_SRC_N     C2       BBAC     8.ch1              8.ch3              BBAC-15
NVM_DIFF_HI    D1       DCDiff   9.ch0              9.ch1              Diff Meter
RLY_K1_MUX     --       Utility  7.bit0             7.bit0             Shared
```

Notes:
- Slots are numbered by physical position in the tester head.
- Multi-site programs duplicate channel assignments per site column.
- Shared resources (e.g., relays) may use the same channel for all
  sites.

---

## 5. Pin Levels

**Header:**

```
DTLevelSheet,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Defines voltage and current levels for pin groups.
UltraFlex levels support spec variable references prefixed with `=`.

**Column headers:**

```
Pin/Group | Seq. | Parameter | Value | Comment
```

**Example rows:**

```
grp_ALL_DIG_PINS              Vih              3.3
grp_ALL_DIG_PINS              Vil              =_vil_val          Spec variable reference
grp_ALL_DIG_PINS              Voh              1.75
grp_ALL_DIG_PINS              Vol              1.55
grp_ALL_DIG_PINS              Ioh              -0.001
grp_ALL_DIG_PINS              Iol              0.001
grp_ALL_DIG_PINS              Vt               1.65
grp_ALL_DIG_PINS              Vcl              =_Vcl_default      Spec variable reference
grp_ALL_DIG_PINS              Vch              =_Vch_default      Spec variable reference
grp_ALL_DIG_PINS              DriverMode       LargeSwing-HiZ
grp_XTAL_PINS                 Vil              0
grp_XTAL_PINS                 Vih              1
grp_XTAL_PINS                 DriverMode       LargeSwing-HiZ
```

Notes:
- Values prefixed with `=` resolve at runtime from the Global Specs
  sheet or DC/AC Specs sheet.
- `DriverMode` controls UltraPin800 output driver behavior
  (`LargeSwing-HiZ`, `SmallSwing-HiZ`, `LargeSwing-Active`,
  `SmallSwing-Active`).

---

## 6. Time Sets

**Header:**

```
DTTimesetBasicSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Defines timing for pattern execution. UltraFlex time sets
support spec variable references for periods and Single / Dual /
Dual-2X timing modes.

**Column headers:**

```
Timing Mode: Single
Time Set Name | Period | Pin/Group | Data Source | Format | Drive On | Drive Data | Drive Return | Drive Off | Compare Mode | Compare Open | Compare Close
```

**Example rows:**

```
ts_dut           =1/_DutClk     grp_ALL_DIG_PINS       NR          NR      0         0           0             =period    Edge          =period*0.95  =period*0.95+1n
ts_flash         =1/_DutClk     grp_ALL_DIG_PINS       NR          NR      0         0           0             =period    Edge          =period*0.95  =period*0.95+1n
ts_flash         =1/_DutClk     DIG_GP25               SBC         SBC     0         =period*0.25 =period*0.75 =period    Edge          =period*0.95  =period*0.95+1n
ts_measfreq      6.25ns         DIG_GP00               NR          NR      0         0           0             6.25ns     Edge          5ns           5.5ns
```

**Key UltraFlex features:**

- Spec variable references in period (`=1/_DutClk` computes 1 / freq).
- SBC (Surround By Complement) drive format for clock generation.
- Edge and Window compare modes.
- Per-pin format overrides within the same time set.
- Expressions using `=period` resolve to the time set period at
  runtime.

---

## 7. Global Specs

**Header:**

```
DTGlobalSpecSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Stores program-wide constants and variables referenced by
other sheets through the `=` prefix.

**Column headers:**

```
Symbol | Job | Value | Comment
```

**Example rows:**

```
HIB_CAL_Status                     0               If > 0 then HIB cal already run
Vcl_default                        -1              Clamp voltage low
Vch_default                        6               Clamp voltage high
Vph_default                        5               Hi-V pin voltage
Iph_default                        0               Hi-V pin current
Tpr_default                        0.000016        Hi-V pin rise time
Vdd_Initial                        3.3             Nominal supply voltage
Vdd_Min                            3.135           Min supply (nom - 5%)
Vdd_Max                            3.465           Max supply (nom + 5%)
Vdd_Stress                         3.63            Stress supply voltage
_DutClk                            26000000        DUT clock frequency (Hz)
_HibChkCalClk                      8000000         HIB checker clock (Hz)
_vil_val                           0               VIL value for spec reference
_Scan_period                       0.0000001       Scan period (100ns)
```

Notes:
- Variables prefixed with `_` are commonly used as spec references in
  levels and timing sheets (e.g., `=_DutClk` in time set period).
- The `Job` column allows per-job overrides of a symbol value.

---

## 8. Job List

**Header:**

```
DTJobListSheet,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Defines selectable test jobs, each binding a flow, pin
map, channel map, and one or more test instance sheets.

**Column headers:**

```
Job Name | Pin Map | Test Instances | Flow | Chan Map | Comment
```

**Example rows:**

```
<PART>_FT_AMB              PinMap_FT     TestInst,TestInst_sp_Digital,TestInst_sp_RADIO  FlowTable_FT_Amb  ChanMap_FT      FT ambient
<PART>_FT_HOT              PinMap_FT     TestInst,TestInst_sp_Digital,TestInst_sp_RADIO  FlowTable_FT_Hot  ChanMap_FT      FT hot
<PART>_FT_COLD             PinMap_FT     TestInst,TestInst_sp_Digital,TestInst_sp_RADIO  FlowTable_FT_Cold ChanMap_FT      FT cold
<PART>_PROBE_AMB           PinMap_Probe  TestInst,TestInst_sp_Digital                    FlowTable_Probe   ChanMap_Probe   Probe ambient
<PART>_Char                PinMap_FT     TestInst,TestInst_sp_Char                       FlowTable_Char    ChanMap_FT      Characterisation
<PART>_Debug_MTF           PinMap_FT     TestInst                                        Flow_Debug_MTF    ChanMap_FT      Debug
```

**Key UltraFlex features:**

- Multiple test instance sheets per job (comma-separated) to support
  sub-programs.
- Separate pin maps for probe versus final test configurations.
- Temperature-specific flow tables for multi-temperature insertion
  programs.

---

## 9. DC Specs and AC Specs

**DC Specs header:**

```
DTDCSpecSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**AC Specs header:**

```
DTACSpecSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

**Purpose:** Define parametric spec values organized by
Category/Selector pairs. Test instances reference these pairs through
the `DC Specs Category` / `DC Specs Selector` and `AC Specs Category`
/ `AC Specs Selector` columns.

**Column headers:**

```
Category | Selector | Symbol | Typ | Min | Max | Comment
```

**Example rows (DC Specs):**

```
Cat0       vddio_nom     VDD_NOM         3.3       3.135     3.465     Nominal IO supply
Cat0       vddio_min     VDD_MIN         3.135     3.0       3.135     Min IO supply
Cat0       vddio_max     VDD_MAX         3.465     3.465     3.63      Max IO supply
```

**Example rows (AC Specs):**

```
Cat0       Sel0          ClkPeriod       38.46n    35n       42n       DUT clock period
Cat0       Sel0          SetupTime       5n        4n        6n        Data setup time
Cat0       Sel_Fast      ClkPeriod       20n       18n       22n       Fast clock period
```

Notes:
- The Category/Selector mechanism allows one program to cover multiple
  operating conditions without duplicating test instances.
- Selectors are referenced by name in the test instances sheet.

---

## 10. Load Board File (.load)

**Purpose:** Configures the tester environment when loading a program.
Sets the active job, enable words, and datalog options.

The format is the same as MicroFlex:

```
Program    \\server\path\to\program.xls
JobName    <PART>_FT_AMB
ChanMap    ChanMap_FT
Env        25C
Part       <PART_NUMBER>

# Enable words
Calibrate  True
Calok      False

# Datalog
DatalogOn           True
DatalogToSTDF       True
STDFDataDirectory   \\server\path\to\stdf\
```

Notes:
- `Calibrate` / `Calok` enable words control which flow sections
  execute on load.
- STDF datalog paths are typically UNC paths to a shared network
  directory.

---

## 11. Sub-Program Structure (UltraFlex-specific)

**Purpose:** UltraFlex programs support sub-programs (`sp_*`
directories) that partition large test programs into independently
maintainable units. Each sub-program can define its own test
instances, flows, timing, levels, and patterns.

**Directory layout:**

```
<program_root>/
  Common/
    src/                        VBT modules shared across all sub-programs
    TestInstances.txt           Main test instances
    Flow_Main.txt               Main flow (calls sub-flows)
    Pinmap.txt                  Main pin map
    ...
  sp_Digital/
    src/VBT_digital.bas         Digital-specific VBT
    TestInst_sp_Digital.txt
    Flow_Digital.txt
    patterns/
  sp_RADIO/
    src/VBT_Radio.bas           Radio-specific VBT
    TestInst_sp_RADIO.txt
    Flow_Radio.txt
    patterns/
  sp_TRIM/
    src/VBT_Trim.bas            Trim-specific VBT
    Flow_Trim.txt
    patterns/
```

Notes:
- Each sub-program has its own `.igxlsp` project file.
- Sub-program test instance sheets are listed in the job list
  `Test Instances` column (comma-separated).
- The main flow calls sub-program flows using the `call` opcode.

---

## 12. Tester Configuration

**Purpose:** Maps instrument boards to physical tester slots for
hardware verification and channel map validation.

**Format:**

```
<slot>.<channel>  <board_type>  <board_serial>
```

**Example rows:**

```
3.0               UltraPin800   UP800-001
5.0               DC-30         DC30-001
8.0               BBAC-15       BBAC-001
9.0               UltraVI80     UVI80-001
12.0              HDVS          HDVS-001
```

Common UltraFlex instrument boards:

| Board         | Purpose                                   |
|---------------|-------------------------------------------|
| UltraPin800   | High-speed digital I/O (800 Mbps)         |
| DC-30         | DC parametric (DCVI) with Kelvin sense    |
| UltraVI80     | High-current DCVI (up to 4 A)             |
| BBAC-15       | Broadband AC source/digitize              |
| HDVS          | High-density VI source                    |
| UltraPAC      | Precision analog capture                  |
| HSD-Ultra     | High-speed digital (1.6 Gbps)             |

---

*End of UltraFlex IG-XL Sheet Format Templates*
