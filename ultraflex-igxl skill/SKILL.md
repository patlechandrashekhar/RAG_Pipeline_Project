---
name: ultraflex-igxl
description: >-
  Comprehensive reference for Teradyne UltraFlex IG-XL test programs:
  platform identification, program structure, VBA/C++ coding conventions,
  VBT API, digital pattern development, and instrument reference.
keywords: >-
  VBA VBT UltraFlex IGXL IG-XL Teradyne tester test-program DCVI PPMU
  SiteDouble SiteLong TheExec TheHdw pattern datalog trim DSP DLL
  digital atp pat pvs DSSC concurrent identify MicroFlex Jaguar
  ultraflex igxl program structure coding vbt api pattern development
---

# UltraFlex IG-XL Comprehensive Reference

## Platform Identification

Inspect cell **A1** of any worksheet in the IG-XL workbook (`.xls`). The sheet header metadata format differs by platform:

### MicroFlex

Cell A1 contains a simple `DFF` tag with a version number:

```
DFF 1.8
```

No platform identifier is present.

### UltraFlex

Cell A1 contains a `DT<SheetType>` tag with embedded key-value metadata that includes `platform=Jaguar` (Teradyne's internal codename for UltraFlex):

```
DTTestInstancesSheet,version=2.4:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
```

The `DT` prefix varies by sheet type (e.g., `DTHomeSheet`, `DTJobListSheet`, `DTChanMap`, `DTFlowtableSheet`, `DTGlobalSpecSheet`, `DTPinMap`).

### Identification Procedure

1. Open the `.xls` workbook with `ViewSpreadsheetStats`.
2. Read the first column name of any sheet.
   - If it starts with `DFF` the program is for **MicroFlex**.
   - If it starts with `DT` and contains `platform=Jaguar` the program is for **UltraFlex**.
3. If the column header is ambiguous, use `ViewSpreadsheetRange` to read cell A1 directly.

## Platform Context

- UltraFlex is an automatic test equipment platform for semiconductor devices, supplied by Teradyne.
- Test programs are authored as Excel workbooks (.xls) containing VBA modules, IGXL ASCII sheets, and test patterns.
- All code targets the UltraFlex/IGXL platform unless otherwise specified.
- Programs may also include C++ custom DSP DLLs and Python offline analysis scripts.

## Reference Material

- For register access via Yoda-generated objects, refer to the Yoda scripting skill.

## IG-XL Workbook Structure

The UltraFlex IG-XL workbook is an Excel `.xls` file. Each sheet has a `DT<SheetType>` header in cell A1 with embedded metadata including `platform=Jaguar`.

### Key Sheets

- **Test Instances** - Defines VBT test instances with parameters (same structure as MicroFlex but with `DT` header)
- **Flow Table** - Sequential test execution flow with opcodes, limits, binning
- **Pin Map** - Logical pin names, types (Analog, I/O, Utility), and pin groups
- **Channel Map** - Maps logical pins to physical tester channels per site
- **Pin Levels** - Voltage levels per pin group (Vil, Vih, Vol, Voh, Vt, Vch, Vcl)
- **Time Sets** - Per-pin timing parameters for each time set
- **Wave Definitions** - Waveform primitives for analog sourcing
- **Global Specs** - System-wide specification symbols
- **Job List** - Jobs combining device variants, test types, and configurations
- **Registers** - DUT register map
- **Board Calibration** - Per-site calibration offsets

### UltraFlex-Specific Sheets

- **Concurrent Test Flow** - Defines concurrent subflows (CSFs) for parallel execution
- **Flow Domains** - Pin resource partitioning for concurrent test
- **Protocol Definition** - Protocol-aware serial interface configurations
- **Characterization Studio** - Integration with characterization tools

## UltraFlex Instrument Boards

### Digital Pin Electronics
- **UltraPin800** - 128 channels, 500 MHz vector rate, 800 Mbps, 3-level drivers (-1V to 6V), HV drive per pin (0-13V), PPMU per pin (2uA-50mA), 4 DSSC source + 16 capture engines, 256M vector memory
- **UltraPin1600** - Higher channel count variant
- **UltraPin4000** - Highest performance variant
- **GigaDig** - High-speed digital
- **HPM (High Performance Memory)** - DDR3/DDR4/GDDR3-5/LPDDR2-3 memory test, up to 4.5 Gbps, uses LMC pattern language

### DC Instruments
- **UltraVI80** - 80 DCVI channels, -2V to +7V, +/-1A (up to 4A merged), 16-bit, voltage spike check, bleeder resistor, pattern-controlled microcodes
- **DC30/DC75** - Legacy DCVI boards (shared with MicroFlex)
- **HDVS** - 48 nonfloating DCVS channels, 0-7V, 1/2/4A merged, 2K capture memory, 64 PSets, internal 1uF decoupling capacitor, pattern-controlled current measurement
- **HexVS** - 6 x 15A DCVS channels, 0-5.5V, mergeable to 90A, pattern-controlled measurement
- **DCIO** - UltraFlex-specific DC I/O instrument
- **HVO2000 (DCVO)** - High voltage output, 16 programmable setups, FVMI/FIMV modes

### Mixed Signal
- **BBAC** - Broadband AC source and capture
- **VHFAC** - Very high frequency AC instruments (source, capture, serial bus, TMU, trigger)
- **AWG6G** - Arbitrary waveform generator
- **UltraPAC** - Precision AC source and capture

### Microwave
- **UW12G/UW24** - Microwave source, receiver, modulated source, multi-tone source, noise source, VNA

### Other
- **PMO** - Precision measurement option
- **SB6G** - 6 Gbps serial bus
- **UltraSerial10G** - 10 Gbps serial

## Tester Configuration

Configuration files map instrument boards to physical slots:

```
Board types: UltraPin800, UltraPin1600, UltraPin4000, GigaDig, HPM,
             UltraVI80, DC-30, DC-75, HDVS, HexVS, HVO2000,
             BBAC-15, VHFAC, AWG6G, UltraPAC, SB6G, UltraSerial10G,
             UW12G, UW24, PMO, SupportBoard, DSP
```

Files: `TesterConfig.txt`, `CurrentConfig.txt`, `CurrentChannelMap.txt`.

## VBA Module Organization

- `Option Explicit` is MANDATORY at the top of every module and class.
- VBT test modules use the `VBT_` prefix: `VBT_ADC.bas`, `VBT_Continuity.bas`, `VBT_PlatformTrim.bas`.
- Support/utility modules use descriptive names without prefix: `Globals.bas`, `Utilities.bas`, `Instruments.bas`, `DSP_Module.bas`.
- Communication modules use protocol names: `SPI.bas`, `SerialWire.bas`, `JTAG.bas`, `Ycomms.bas`.
- Classes use descriptive names: `CDll.cls`, `DUTconditions.cls`, `clsTimer.cls`.
- `RunVBT.bas` is auto-generated and MUST NOT be edited manually.
- Each module should have an owner attribution comment near the top: `'''Owner: <name>`.
- Sub-program directories use the `sp_` prefix: `sp_Digital/`, `sp_RADIO/`, `sp_TRIM/`, `sp_Flash/`.

## Naming Conventions

### Variable Prefixes

- `g_` for global variables: `g_CurrentTestCondition`, `g_DIE_Temp`, `g_PROGRAM`
- `c_` for constants: `c_NomTC`, `c_DiffHiRng`, `c_BMANN`
- `m_` for private object member variables
- No prefix for module-level private variables or local variables

### Variable Naming

- Use descriptive camelCase or PascalCase names.
- Site-aware variables MUST be declared with explicit IGXL types using `As New`:
  - `Dim trimCode As New SiteLong`
  - `Dim measurementResult As New SiteDouble`
  - `Dim devicePassed As New SiteBoolean`
  - `Dim perPinVoltage As New PinListData`
  - `Dim capturedWaveform As New DSPWave`

### Naming Examples

Good:
- `settleDelay`, `trimCode`, `measurementResult`, `devicePassed`
- `TF_SetupDUT()`, `TF_Continuity()`, `TF_SetLevels()`
- `c_NominalSupply`, `g_SiliconRev`, `g_GlobalVarReset`

Avoid:
- `v1`, `tmp`, `x`, `res`
- Single-character loop variables other than `i`, `j`, `k`

## VBT Function Structure

Every VBT function follows this canonical skeleton:

```vba
Public Function MyTest(pinList As PinList, pat As Pattern) As Long
    On Error GoTo errHandler

    Dim measResults As New PinListData

    ' --- Setup hardware ---
    ' --- Perform measurement ---
    ' --- Evaluate results via TestLimit ---
    ' --- Cleanup ---

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
```

Rules:
- Top-level test functions callable from the flow MUST use the `TF_` prefix and return `Long`:
  ```vba
  Public Function TF_MeasureSupplyCurrent() As Long
  ```
- Helper/worker functions omit the `TF_` prefix:
  ```vba
  Public Function MeasureAdcChannel(channelIndex As Long) As Double
  ```
- Error handler uses `If AbortTest Then Exit Function Else Resume Next`
- Function must be in a module whose name starts with `VBT` (e.g., `VBT_Module`)
- Names must NOT resemble Excel cell references (avoid up to 3 letters + up to 7 numbers)
- Support functions should be `Private` or placed in non-VBT modules
- Parameters: `Single`, `Double`, `Long`, `Integer`, `String`, `Boolean`, `Enum`, `Pattern`, `PatternSet`, `PinList`, `InterposeName`
- User-defined `Enum` types appear as drop-down lists in the Instance Editor
- AbortTest (from Template.xla) logs errors and determines abort vs. continue
- For additional error logging in production:
  ```vba
  errHandler:
      TheExec.Datalog.WriteComment "Error in TF_ExampleTest: " & _
          Err.Number & " " & Err.Description
      If AbortTest Then Exit Function Else Resume Next
  ```

## Core Objects

### TheExec (Execution Control)

```vba
TheExec.Flow.TestLimit ...           ' Datalog and limit checking
TheExec.Flow.TestNumber = N          ' Set test number for datalogging
TheExec.Flow.TestLimitIndex = N      ' Jump to specific limit row
TheExec.Sites                        ' Site iteration (Selected, Active, Existing)
TheExec.Sites.Existing.Count         ' Total sites in channel map
TheExec.Datalog.WriteComment "msg"   ' Write to datalog
TheExec.Datalog.SetDynamicTestName   ' Dynamically set test name
TheExec.DataManager.DecomposePinList "group", pinArray, numPins
TheExec.TesterMode                   ' testModeOnline / testModeOffline
TheExec.RunMode                      ' runModeProduction / runModeDebug
TheExec.CurrentJob                   ' Identify current job
TheExec.CurrentChanMap               ' Active channel map name
TheExec.EnableWord("name") = True    ' Enable word control
TheExec.ExecutionCount               ' Detect first run
TheExec.AddOutput "text"             ' Write to Test Program Output window
TheExec.ErrorLogMessage "msg"        ' Error logging (causes flow to fail the test)
```

### TheHdw (Hardware Access)

Top-level object for all tester hardware. Supports two access patterns:

```vba
' By-instrument (standard)
TheHdw.DCVI.Pins("p0").CurrentRange = 0.002

' By-pin (equivalent)
TheHdw.Pins("p0").DCVI.CurrentRange = 0.002
```

Complete instrument interfaces on UltraFlex:

```vba
TheHdw.DCVI.Pins("pinList")        ' DC Voltage/Current (DC30/DC75/UltraVI80)
TheHdw.DCVS.Pins("pinList")        ' DC Voltage Supply (HexVS/HDVS)
TheHdw.DCIO.Pins("pinList")        ' DC I/O instrument (UltraFlex-specific)
TheHdw.Digital                       ' Digital subsystem (UltraPin800/GigaDig/HPM)
TheHdw.PPMU.Pins("pinList")        ' Per-Pin Measurement Unit
TheHdw.BBACSource.Pins("pinList")   ' Broadband AC Source
TheHdw.BBACCapture.Pins("pinList")  ' Broadband AC Capture
TheHdw.DCTime.Pins("pinList")       ' DC Time measurement
TheHdw.DCVO.Pins("pinList")         ' DC Voltage Output (HVO2000)
TheHdw.DCDiffMeter.Pins("pinList")  ' DC Differential Meter
TheHdw.DSSC("pinName")              ' Digital Source/Capture (SPI)
TheHdw.DSP                           ' DSP subsystem
TheHdw.Utility.Pins("pinList")      ' Utility bits (relays)
TheHdw.PinLevels.Pins("pinList")    ' Pin level control
TheHdw.Alarm                         ' Hardware alarm management
TheHdw.FlowDomain                    ' Concurrent test flow domain control
TheHdw.GPIB                          ' IEEE-488 bus control
TheHdw.Protocol                      ' Protocol-aware serial interfaces
TheHdw.SyncPanel                     ' Synchronization panel
TheHdw.TDR                           ' Time domain reflectometry
TheHdw.PLMeter.Pins("pinList")      ' Power level meter
' VHF AC instruments
TheHdw.VHFACCapture.Pins("pinList")
TheHdw.VHFACSource.Pins("pinList")
TheHdw.VHFACTime.Pins("pinList")
TheHdw.VHFACSerialBus
TheHdw.VHFACTrigControl
' Microwave instruments
TheHdw.MW                            ' Microwave subsystem
TheHdw.MWSource.Pins("pinList")
TheHdw.MWReceiver.Pins("pinList")
TheHdw.MWModulatedSource.Pins("pinList")
TheHdw.MWMultiToneSource.Pins("pinList")
TheHdw.MWNoiseSource.Pins("pinList")
' Other
TheHdw.GigaDig                       ' High-speed digital
TheHdw.AVI64                         ' 64-channel analog
TheHdw.Serial                        ' Serial communication
TheHdw.Wait timeInSeconds             ' Timed delay
```

## DCVI API (DC Voltage/Current Instrument)

Works with DC30, DC75, and UltraVI80 boards.

### Modes
```vba
TheHdw.DCVI.Pins("pin").Mode = tlDCVIModeVoltage       ' Force voltage
TheHdw.DCVI.Pins("pin").Mode = tlDCVIModeCurrent       ' Force current
TheHdw.DCVI.Pins("pin").Mode = tlDCVIModeHighImpedance ' High-Z (measure only)
TheHdw.DCVI.Pins("pin").Mode = tlDCVIModeHighRegulation ' High regulation
```

### Voltage and Current
```vba
TheHdw.DCVI.Pins("pin").Voltage = 3.3
TheHdw.DCVI.Pins("pin").Current = 0.001
TheHdw.DCVI.Pins("pin").SetVoltageAndRange 3.3, 5
TheHdw.DCVI.Pins("pin").SetCurrentAndRange 20 * mA, 20 * mA
```

### Ranges
```vba
TheHdw.DCVI.Pins("pin").VoltageRange = 5
TheHdw.DCVI.Pins("pin").VoltageRange.Autorange = False
TheHdw.DCVI.Pins("pin").CurrentRange = 0.002
TheHdw.DCVI.Pins("pin").ComplianceRange(tlDCVIComplianceBoth) = 30
```

### Connection and Gate
```vba
TheHdw.DCVI.Pins("pin").Connect                        ' default: force+sense+guard
TheHdw.DCVI.Pins("pin").Connect tlDCVIConnectHighSense
TheHdw.DCVI.Pins("pin").Disconnect
TheHdw.DCVI.Pins("pin").Gate = True                    ' enable output
' Gate modes: True (on), False (off, V=0V)
' UltraVI80 also supports: tlDCVIGateOn, tlDCVIGateOff, tlDCVIGateOffHiZ
```

### Meter
```vba
TheHdw.DCVI.Pins("pin").Meter.Mode = tlDCVIMeterCurrent
TheHdw.DCVI.Pins("pin").Meter.Filter.Bypass = False
TheHdw.DCVI.Pins("pin").Meter.Filter.Value = 500
Dim results As New PinListData
results = TheHdw.DCVI.Pins("pins").Meter.Read(tlStrobe, sampleSize, sampleRate)
```

### Alarms
```vba
TheHdw.DCVI.Pins("pin").Alarm(tlDCVIAlarmAll) = tlAlarmContinue
' Behaviors: tlAlarmForceFail (default), tlAlarmForceBin, tlAlarmOff, tlAlarmContinue, tlAlarmDefault
```

### PSets (Parameter Sets)
```vba
TheHdw.DCVI.Pins("pin").PSets.Add "MyPSet"
TheHdw.DCVI.Pins("pin").PSets.Item("MyPSet").Set _
    Mode:=tlDCVIModeVoltage, Voltage:=3.3, VoltageRange:=5, _
    Current:=0.1, CurrentRange:=0.2, MeterMode:=tlDCVIMeterCurrent
TheHdw.DCVI.Pins("pin").PSets.Item("MyPSet").Apply

' Apply from pattern microcode (faster, no VBT overhead)
' Use PSet microcode opcode in pattern
```

Benefits: Faster than individual property programming, atomic state changes, pattern-controllable. Stored in instrument board memory.

### UltraVI80-Specific Properties
```vba
TheHdw.DCVI.Pins("pin").NominalBandwidth = 500
TheHdw.DCVI.Pins("pin").CrossOverType = tlDCVICrossOverTypeAuto
TheHdw.DCVI.Pins("pin").LocalKelvin(tlDCVILocalKelvinLow) = False
TheHdw.DCVI.Pins("pin").LimiterEnabled = True
TheHdw.DCVI.Pins("pin").CurrentLimitMode = tlDCVICurrentLimitModeClamp
TheHdw.DCVI.Pins("pin").BleederResistor = True
' Pattern-controlled microcodes: Gate_On, Gate_Off, Gate_Off_HiZ, Strobe
' Voltage spike check function monitor per channel
```

## DCVS API (DC Voltage Supply)

Works with HexVS and HDVS boards.

```vba
TheHdw.DCVS.Pins("pin").Connect tlDCVSConnectDefault    ' force+sense+guard
TheHdw.DCVS.Pins("pin").Voltage.Main = 3.3              ' main voltage (0-5.5V HexVS)
TheHdw.DCVS.Pins("pin").Voltage.Alt = 3.6               ' alternate voltage
TheHdw.DCVS.Pins("pin").Voltage.Output = tlDCVSVoltageMain  ' select main/alt
TheHdw.DCVS.Pins("pin").CurrentRange = 15               ' HexVS: 15/30/60/90A merged
TheHdw.DCVS.Pins("pin").CurrentLimit.Source.FoldLimit.Level = 1.0
TheHdw.DCVS.Pins("pin").CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
TheHdw.DCVS.Pins("pin").Meter.Mode = tlDCVSMeterCurrent
TheHdw.DCVS.Pins("pin").Meter.Filter.Bypass = True
results = TheHdw.DCVS.Pins("pin").Meter.Read(tlStrobe, 20, 25000)
TheHdw.DCVS.Pins("pin").BandwidthSetting = 4            ' 0-7 for HexVS
TheHdw.DCVS.Pins("pin").BleederResistor = True
TheHdw.DCVS.Pins("pin").Alarm(tlDCVSAlarmAll) = tlAlarmContinue
TheHdw.DCVS.Pins("pin").Disconnect
```

HDVS-specific: 48 channels per board, 0-7V, 1/2/4A merged, pattern-controlled measurement, 2K capture memory, 64 PSets, internal 1uF decoupling capacitor.

## DCIO API (UltraFlex-Specific)

```vba
TheHdw.DCIO.Pins("pin").Mode = tlDCIOModeVoltage
TheHdw.DCIO.Pins("pin").Voltage = 3.3
TheHdw.DCIO.Pins("pin").Current = 0.001
TheHdw.DCIO.Pins("pin").Connect
TheHdw.DCIO.Pins("pin").Meter.Read(tlStrobe, sampleSize)
```

## Digital Subsystem API

Works with UltraPin800, GigaDig, and HPM boards.

```vba
' Apply levels and timing
TheHdw.Digital.ApplyLevelsTiming ConnectAllPins:=True, _
    loadLevels:=True, loadTiming:=True, RelayMode:=tlUnpowered

' Pin connection
TheHdw.Digital.ConnectPins "DUT_DIG"
TheHdw.Digital.DisconnectPins "pins"

' Pattern execution
TheHdw.Digital.Patterns.Pat("pattern.pat").Run ""
TheHdw.Digital.Patterns.Pat("pattern.pat").Run "StartLabel"
TheHdw.Digital.Patterns.Pat("pattern.pat").Start ""
TheHdw.Digital.Patgen.HaltWait

' Check pass/fail per site
For Each site In TheExec.Sites.Active
    passed = TheHdw.Digital.Patgen.PatternBurstPassed(site)
Next site

' Channel failure check
failed = TheHdw.Digital.ChannelFailed(channelNum)

' Pin level modification
origLevel = TheHdw.PinLevels.Pins("pin").ReadPinLevels(chVih)
Call TheHdw.PinLevels.Pins("pin").ModifyLevel(chVih, newValue)

' High-voltage powered mode (UltraFlex-specific)
TheHdw.Digital.ForcedHVPoweredMode = True  ' enable hot switching HV relays
' Resets to False at end of job unless explicitly set

' UserDib exclude pins from calibration
Call TheHdw.Digital.UserDibExcludePins("a1, a2, a3", 5E-09)

' Pattern jump (UltraPin800)
TheHdw.Digital.Patgen.SetJump "label"
TheHdw.Digital.Patgen.Continue

' Modify vector data per-site
TheHdw.Digital.Patterns.Pat("pattern.pat").ModifyVectorOperand "", position, value
TheHdw.Digital.Patterns.Pat("pattern.pat").ModifyPinVectorBlockDataSite( _
    "", offset, "SPI_MOSI", binaryString, siteNumber)

' Load all patterns
TheExec.Patterns.LoadAll
```

## PPMU API (Per-Pin Measurement Unit)

Built into each digital board (UltraPin800: 2uA-50mA; HPM: 20uA-32mA).

```vba
TheHdw.PPMU.Pins("pins").ForceV 0, 2 * mA
TheHdw.PPMU.Pins("pins").ForceI 0.0001
TheHdw.PPMU.Pins("pins").MeasureI                       ' force V, measure I
TheHdw.PPMU.Pins("pins").MeasureV                       ' force I, measure V
TheHdw.PPMU.Pins("pins").SetClamps -1, 6                ' low, high voltage clamps
TheHdw.PPMU.Pins("pins").Connect
TheHdw.PPMU.Pins("pins").Gate = tlOn
results = TheHdw.PPMU.Pins("pins").Read(tlPPMUReadMeasurements, sampleCount)
TheHdw.PPMU.Pins("pins").Disconnect
```

## BBAC API (Broadband AC Source/Capture)

### Source
```vba
TheHdw.BBACSource.Pins("pin").Amplitude.Value = 3       ' Volts peak differential
TheHdw.BBACSource.Pins("pin").SampleRate.Value = 500000
TheHdw.BBACSource.Pins("pin").VoltageRange.Value = 5
TheHdw.BBACSource.Pins("pin").VoltageRange.Autorange = True
TheHdw.BBACSource.Pins("pin").CommonModeVoltage.Value = 1.5
TheHdw.BBACSource.Pins("pin").Connect                   ' default from channel map
TheHdw.BBACSource.Pins("pin").Connect tlBBACSourceFromREF, tlBBACSourceToDUT

' Signal-level control
With TheHdw.BBACSource.Pins("pin").Signals.Item("sig")
    .Amplitude = 3
    .ConnectionType = tlBBACSourceConnectionTypeDifferential
    .CommonMode.Enable = True
    .CommonMode.Value = 0.5
    .Start tlStartImmediately
End With
```

### Capture
```vba
TheHdw.BBACCapture.Pins("pin").SampleRate.Value = 500000
TheHdw.BBACCapture.Pins("pin").SampleSize.Value = 2048
TheHdw.BBACCapture.Pins("pin").VoltageRange.Value = 5
TheHdw.BBACCapture.Pins("pin").Offset.Value = 2         ' DC offset removal
TheHdw.BBACCapture.Pins("pin").Connect

' Retrieve captured data
Dim dspData As New PinListData
Set dspData = TheHdw.BBACCapture.Pins("pin").Signals.Item("sig").DSPWave
```

## DCTime API

```vba
TheHdw.DCTime.Pins("pin").Connect
TheHdw.DCTime.Pins("pin").Mode = tlDCTimeModeStamper
TheHdw.DCTime.Pins("pin").Measurement.Frequency.SetFrontEnd _
    10, tlDCTimeImpedanceHiZ, 5, tlDCTimeHysteresisOff
TheHdw.DCTime.Pins("pin").Measurement.Frequency.Start
' Measurement types: Frequency, Period, DutyCycle, PulseWidth, RiseTime, FallTime, PinToPinDelay
```

## DCVO API (HVO2000)

```vba
' Configure setup
TheHdw.DCVO.Pins("pin").Setups(tlDCVOSetupsEntry00).Set _
    Mode:=tlDCVOModeForceVMeasureI, Voltage:=5, Current:=0.1, _
    PulseWidth:=0.001, forceVRange:=10, meteringRange:=0.1, _
    highLimit:=0.15, lowLimit:=0, avgTime:=0.0001, _
    avgTimeEnab:=True, gotoNextSetupEnab:=False

' Execute and read
Call TheHdw.DCVO.Pins("pin").Setups(tlDCVOSetupsEntry00).Start
While TheHdw.DCVO.Pins("pin").IsRunning
    TheHdw.Wait 1 * ms
Wend
Dim result As New PinListData
result = TheHdw.DCVO.Pins("pin").Setups(tlDCVOSetupsEntry00).Meter.Read(tlDCVOMeterFormatResults)
```

## DSSC API (Digital Source/Capture)

UltraPin800: 4 source engines (32-bit) + 16 capture engines (8-bit, mergeable to 32-bit) per board.

### Source (write data to DIN)
```vba
TheExec.WaveDefinitions.CreateWaveDefinition signalName, wordData, True
With TheHdw.DSSC("DIN").Pattern(patName).Source.Signals
    .Add signalName
    .DefaultSignal = signalName
End With
With TheHdw.DSSC("DIN").Pattern(patName).Source.Signals(signalName)
    .WaveDefinitionName = signalName
    .SampleSize = totalNumWords
    .Amplitude = 1
    .LoadSamples
End With
```

### Capture (read data from DOUT)
```vba
With TheHdw.DSSC.Pins("DOUT").Pattern(patName).Capture
    .Signals.Add signalName
    .Signals(signalName).SampleSize = numberOfBytes
    .Signals(signalName).LoadSettings
End With

' Retrieve captured data
Dim captureData As DSPWave
Set captureData = TheHdw.DSSC.Pins("DOUT").Pattern(patName).Capture.Signals(signalName).DSPWave
value = captureData.Element(0)
```

### Pattern Markers for DSSC
- Use `stv` microcode opcode to mark vectors for capture
- Use Edge Capture compare mode on capture pins
- Source data loaded from wave definitions before pattern run

## PinListData

Multisite, multipin data container returned by all instrument measurements.

```vba
Dim measResults As New PinListData
measResults = TheHdw.DCVI.Pins("pins").Meter.Read(tlStrobe, 10, 1000)

' Access per-pin value (inside site loop)
For Each site In TheExec.Sites
    val = measResults.Pins("pinName").Value
Next site

' Built-in math (returns new PinListData)
Dim scaled As PinListData
Set scaled = measResults.Math.Divide(2)
Set scaled = measResults.Math.Add(5)

' Pass to TestLimit
TheExec.Flow.TestLimit resultVal:=measResults, unit:=unitVolt, ForceResults:=tlForceFlow
```

## SiteVariant Types

Site-aware scalar types. Always use `Dim As New`.

```vba
Dim result As New SiteDouble
Dim count As New SiteLong
Dim flag As New SiteBoolean

' Math operations (deferred evaluation for background DSP)
result = result.Add(1).Multiply(2.4)   ' cascading supported
result.Subtract x
result.Divide y
result.Truncate

' Comparison (returns SiteBoolean)
Dim cond As SiteBoolean
cond = result.Compare(GreaterThan, 5#)
If cond.Any(True) Then ...

' IMPORTANT: Deferred evaluation
' Math operations are queued and evaluated only when results are accessed
' (e.g., by TestLimit). Cascading is supported.
' Memory leak warning: globally declared SiteVariants with deferred math
' assigned to themselves that are never evaluated can leak memory.
' Solution: Set MyVar = Nothing in OnProgramEnded.
```

## TestLimit Method

```vba
TheExec.Flow.TestLimit _
    resultVal:=value, _          ' PinListData, SiteDouble, SiteLong, Double, Long, DSPWave.Element()
    lowVal:=0#, _                ' optional (prefer flow table limits)
    hiVal:=1#, _                 ' optional
    lowCompareSign:=tlSignGreaterEqual, _  ' tlSignEqual/Greater/GreaterEqual/Less/LessEqual/None/NotEqual
    highCompareSign:=tlSignLessEqual, _
    ScaleType:=scaleNoScaling, _ ' scaleNone (auto), scaleNoScaling, scaleMilli, scaleMicro, etc.
    unit:=unitVolt, _            ' unitVolt, unitAmp, unitHz, unitTime, unitDb, unitNone, unitCustom
    customUnit:="mA", _          ' required when unit:=unitCustom
    formatStr:="%6.4f", _        ' ANSI C printf format
    ForceResults:=tlForceFlow, _ ' tlForceNone, tlForcePass, tlForceFail, tlForceFlow, tlForceFlowPass, tlForceFlowFail
    forceVal:=3.3, _             ' forced condition value
    forceunit:=unitVolt, _
    TName:="TestName", _         ' override test name
    TNum:=1000, _                ' override test number (required for deferred limits)
    pinName:="pin", _
    compareMode:=CompareAverage  ' CompareAverage or CompareEachSample (for arrays)
```

TName resolution order: 1) TestLimit TName parameter, 2) Flow Table TName column, 3) Test instance name.

TestLimit also closes alarm window and reports any alarms.

ALWAYS use `ForceResults:=tlForceFlow` to ensure results are logged. For dimensionless values (trim codes, counts), use `ScaleType:=scaleNone, forceunit:=unitNone`.

## Site Loops

```vba
' Standard site iteration
Dim site As Variant
For Each site In TheExec.Sites
    val = measResults.Pins("pin").Value   ' site-aware
Next site

' Selectively run on a subset of sites
TheExec.Sites.Selected = siteMask
' ... operations on selected sites ...
TheExec.Sites.Selected = True  ' restore all active

' IMPORTANT: Site-specific SiteVariant accessed outside site loop = runtime error
' Site-uniform SiteVariant works anywhere like a scalar
```

## Unit Multiplier Constants

`ms`, `us`, `s`, `mA`, `uA`, `nA`, `mV`
Usage: `5 * ms`, `200 * uA`, `20 * mA`

## Pattern File Formats

### .atp (ASCII Text Pattern)

Human-readable source format. Structure:

```
opcode_mode = single;
digital_inst = hsd;
import tset spi_timing;
import subr Write_header;

vector ($tset, SCLK, DIN, CSB, DOUT)
{
start_label WriteSPI:
> spi_timing  0 0 1 X ;
> spi_timing  0 0 0 X ;     // CSB low
> spi_timing  0 1 0 X ;     // data bit
> spi_timing  1 1 0 X ;     // clock rising edge
> spi_timing  0 0 1 X ;     // CSB high
halt;
}
```

Key header statements:
- `opcode_mode = single` - One opcode per vector (also `dual`, `dual_2x`)
- `digital_inst = hsd` - Target instrument (hsd, ultraflex, hpm)
- `import tset <name>` - Import time set definitions
- `import subr <name>` - Import subroutine definitions
- `check_vm_min_size = <n>` - Minimum vector memory size check
- `vector ($tset, pin1, pin2, ...)` - Pin declaration block

### .pat (Compiled Binary Pattern)

Binary format compiled from .atp sources by the IG-XL pattern compiler.

### .pvs (Pattern Vector Set)

Associates a pattern file with its pin/timing configuration.

## UltraFlex Timing Modes

### Single Mode
- 32 time sets available
- One opcode per vector cycle
- Standard mode for most applications

### Dual Mode
- 16 time sets available
- Two opcodes per vector cycle (opcode + data)
- Higher throughput for data-intensive patterns

### Dual-2X Mode
- 16 time sets available
- Two vectors per cycle at double rate
- 256M vector memory
- Drive formats have 2X variants (NR_2X, RH_2X, etc.)

## Vector Notation

| Symbol | Meaning |
|--------|---------|
| `0`    | Drive low |
| `1`    | Drive high |
| `X`    | Don't care (no drive, no compare) |
| `H`    | Expect high |
| `L`    | Expect low |
| `C`    | Capture data (Edge Capture mode) |
| `V`    | Capture and verify |
| `Z`    | High impedance |
| `T`    | Terminate |
| `M`    | Midband (three-level driver) |

Drive formats (UltraPin800):
- `NR` (Non-Return), `RH` (Return-High), `RL` (Return-Low)
- `SBC` (Surround By Complement), `SBH` (Surround By High), `SBL` (Surround By Low)
- `STAY` (maintain previous state)
- 2X variants: `NR_2X`, `RH_2X`, `RL_2X`, etc.

Compare modes:
- `Edge` - Single strobe comparison
- `Window` - Dual strobe window comparison (single mode only)
- `Edge Capture` - Capture data on strobe (for DSSC)
- `Off` - No comparison

## Pattern Microcode (Opcodes)

### Flow Control
```
halt                          ; stop execution
repeat <count>                ; repeat next vector N times
loop <label>, <count>         ; loop to label N times
end_loop                      ; end of loop block
jump <label>                  ; unconditional jump
call <subroutine>             ; call subroutine
return                        ; return from subroutine
match <label>, <count>        ; conditional loop (wait for match)
match_vm <label>, <count>     ; conditional branch on fail/pass flags
branch_expr <label>           ; conditional jump based on expression
```

### Counter Stack (UltraFlex-specific)
```
set_msb <value>               ; set most significant bits
set <value>                   ; set counter value
push_loop                     ; push loop counter
pop_loop                      ; pop and decrement loop counter
```

### DSSC Control
```
stv                           ; store this vector (capture marker)
mask                          ; mask comparison for this vector
```

### CPU Flag Handshaking
```
cpuA, cpuB, cpuC, cpuD       ; CPU-controlled flags for VBT-pattern synchronization
```

### Pipeline Control (UltraFlex-specific)
```
pipe_minus                    ; pipeline flush (255 cycles on UltraPin800)
```

## Pattern Types

### Vector Memory (vm_vector)
Standard pattern vectors stored in vector memory. Default type.

### Subroutine Memory (srm_vector)
Subroutine vectors stored in separate subroutine memory for reuse.

## Multiple Time Domains (MTD)

UltraFlex supports different timing domains on separate digital boards. Each time domain has its own clock and timing set.

```vba
' Access time domain configuration
TheHdw.Digital.TimeDomains()

' Pattern header for MTD
' digital_inst = hsd;
' Separate time set imports per domain
```

Restrictions: HPM does not support MTD.

## Keepalive Patterns

Maintain continuous clocking and instrument states (levels, timing, power) across pattern bursts:

```vba
' Erase existing keepalive RAM
Call TheHdw.Digital.keepAlive.Pins("DUT_DIG").EraseRAM

' Program keepalive vector
Call TheHdw.Digital.keepAlive.Pins("CLK").SetRAM(vectorGroup, "1")
TheHdw.Digital.keepAlive.count = loopLength
```

## Pattern Threading (UltraFlex-specific)

Allows different patterns to run on different sites simultaneously.

```vba
' Patterns per site using SetJump
TheHdw.Digital.Patgen.SetJump "site_specific_label"
```

## HRAM (History RAM)

Configurable capture for pattern debug:

```vba
' Configure capture type
Call TheHdw.Digital.HRAM.SetCapture(captType, compressRepeats)
' captType: captAll, captFail, captSTV, captFailSTV, captFirstModFail, captPass, captFirstModPass

' Configure trigger
Call TheHdw.Digital.HRAM.SetTrigger(trigType, waitForEvent, preTrigCycles, stopOnFull)
' trigType: trigFirst, trigFail, trigSTV

' Read results
Call TheHdw.Digital.HRAM.GetCapture(retCapt, retCompress)
Call TheHdw.Digital.HRAM.GetTrigger(retTrig, retWait, retCycles, retStop)
```

## HPM Pattern Language (LMC)

HPM (High Performance Memory) boards use Logical MicroCode (LMC) instead of standard Pattern Language for memory-specific test algorithms. LMC supports algorithmic address/data generation for DDR/GDDR/LPDDR memory protocols.

## UltraPin800 Specifications

- 128 single-ended channels (64 differential pairs) per board
- 500 MHz vector rate (2 ns period), 800 Mbps data rate
- 250 MHz opcode rate
- 256M vector memory in dual-2X mode
- Three-level drivers: VIL, VIH, VT from -1V to 6V (1 mV resolution)
- High-voltage drive: 0-13V at +/-10 mA per pin
- Programmable current loads: +/-20 mA per channel
- Edge and window strobe; +/-150 ps EPA
- Enhanced timing accuracy: +/-80 ps within 64-channel group
- Differential drive/compare with differential comparators
- PPMU per pin (2 uA to 50 mA)
- Frequency counter per channel
- 255-cycle pipeline depth
- 10-64 SCAN chains supported
- 500 Mbps Memory Test Option (MTO)

## Concurrent Test

UltraFlex supports executing tests on multiple device functional blocks simultaneously.

### Concepts
- **Concurrent Subflows (CSFs)** - Run in parallel from calling flow using `concurrent` opcode
- **Flow Domains** - Partition pin resources among concurrent subflows
- **Concurrent Test Efficiency (CTE)** - Metric comparing concurrent vs. serial execution time
- **Site Symmetrical/Asymmetrical** - Execution modes across sites

### Programming
```vba
' Flow domain control
TheHdw.FlowDomain  ' Access concurrent test flow domain

' Pin resource partitioning required between concurrent subflows
' Synchronous pattern start across concurrent subflows supported
```

### Concurrent Profiler
Tool that measures foreground/background execution time to optimize concurrent flow creation.

## Exec Interpose Functions (Program Lifecycle)

Lifecycle hooks in `Exec_IP_Module.bas`:

- **OnTesterInitialized** - Early initialization before TheExec established
- **OnProgramLoaded** - Add references, set enable words, configure sort order
- **OnProgramValidated** - Set starting sites, DIB power on, load patterns, configure datalog
- **OnProgramStarted** - Per-device init (registers, instruments, global vars)
- **OnProgramEnded** - Power down DUT, evaluate calibration, set enable words
- **OnPreShutDownSite / OnPostShutDownSite** - Per-site power-down

## DSP Subsystem

Built-in DSP engine for analog/mixed-signal measurements:
- Signal creation (AWG waveforms)
- Signal analysis (FFT, THD, SNR, SFDR, etc.)
- Background DSP with deferred evaluation
- Custom DSP DLL support
- Integration with DCVI capture/source and DSSC

## RunVBT Module

Auto-generated wrappers. MUST NOT be edited manually.

```vba
Public Function TestName__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Then On Error GoTo errpt
    TestName__ = VBAProject.VBT_Module.TestName(v(0), CDbl(v(1)))
    Exit Function
errpt:
    HandleUntrappedError
End Function
```

## VBT Libraries

Shared VBT code in referenced `.xla` library files. Local workbook functions take precedence over identically named library functions.

## Register Access

### Yoda-Generated Register Objects

Register read/write uses Yoda-generated objects with `regWrite`/`regRead`, `bfWrite`/`bfRead`:

```vba
regWrite RM_gpio.GP0DS, &H0&, DUT, NOVERIF, True
Dim registerValue As Long
registerValue = regRead(RM_adc.ADCCON, DUT)
bfWrite BF_hpbg.TSTEN, 1, DUT, NOVERIF, True
Dim bitfieldValue As Long
bitfieldValue = bfRead(BF_anaplt.ALDOVTRIM, DUT)
```

### JTAG Firmware Register Access

Memory read/write via JTAG firmware class:

```vba
Call theJtagFw.JtagMemRead(RM_WDT_WDT.CCOUNT, memRd(), 1)
Call theJtagFw.jtag_mem_write(RM_PMU_PMG.PWRKEY, &H4859&)
```

### Shared Memory Communication

DUT communication via shared memory addresses:

```vba
SW_MemWrite COMM_CMD_ADDR, CMD_ADC_PAR_DATA_INT, True
SW_MemRead COMM_BUFFER_ADCRes2, adcBusyCount
```

## Enums and User-Defined Types

Use `Public Enum` for configuration options and state definitions:

```vba
Public Enum MuxSelectEnum
    MUX_DISABLED = 0
    MUX_CHANNEL_A = 1
    MUX_CHANNEL_B = 2
End Enum
```

Use `Public Type` for grouping related test condition data:

```vba
Public Type TestCondition
    supplyVoltage As Double
    temperature As Double
    clockFrequency As Double
End Type
```

## Conditional Compilation

Conditional compile constants are set in Tools -> Project Properties -> Conditional Compile Arguments.

Use bitmask patterns for program variant selection:

```vba
#If (PGM_MASK And PROBE) Then
    ' Probe-specific code
#ElseIf (PGM_MASK And FT) Then
    ' Final test code
#End If
```

Debug output:

```vba
#If (SYS_DEBUG = 1) Then
    Debug.Print "Debug: " & variableName
#End If
```

## C++ Custom DSP DLL Conventions

- Export functions using the `TL_DSP_LIB` macro from `teradyne/tl_dspcustomlibincludes.h`.
- First parameter is always `long *error_code` for status output.
- Use `VARIANT*` for SAFEARRAY input/output arrays.
- Return `S_OK`, `E_POINTER`, or `E_INVALIDARG` as HRESULT.
- Validate all VARIANT types before use (check `.vt` field).
- Free all dynamically allocated memory explicitly.
- Expose version info via a `get_version()` function.
- Wrap DLL access in a VBA class (e.g., `CDll.cls`) using `Declare Function` for `LoadLibrary`/`FreeLibrary`.

## VBT Code Profiler

```vba
' Insert timer marks for profiling
Call ProfileMark("mark-name")
' Enable via Run Options > Detailed Execution Time > VBT Profiler
```

Legacy timer approach:

```vba
theTimer.VBTFunctionStart
' ... test code ...
theTimer.VBTFunctionFinish
```

## Key Differences from MicroFlex

| Feature | MicroFlex | UltraFlex |
|---------|-----------|-----------|
| Workbook header | `DFF 1.8` | `DT...platform=Jaguar` |
| Digital PE | HSD-200, HVD-1 | UltraPin800/1600/4000, GigaDig, HPM |
| DCVI | DC-30, DC-75, DC-90 | DC-30, DC-75, UltraVI80 |
| DCVS | HexVS | HexVS, HDVS |
| Additional DC | - | DCIO |
| Concurrent Test | Not supported | Supported (CSFs, flow domains) |
| Multiple Time Domains | Not supported | Supported |
| Background DSP | Not supported | Supported (deferred evaluation) |
| Microwave | Limited (MWSrc/Rec) | Full MW subsystem (UW12G/UW24) |
| VHFAC | Basic | Full subsystem (source, capture, TMU, serial bus) |
| Protocol Aware | Not supported | Supported (Protocol Studio) |
| Memory Test | Basic | HPM with LMC pattern language |
| By-pin syntax | Not supported | `TheHdw.Pins("p0").DCVI` equivalent |
| HV powered mode | Not available | `ForcedHVPoweredMode` |
| VBT Code Profiler | Basic timer | Full profiler with `ProfileMark()` |

## Prohibited Patterns

- Do not edit `RunVBT.bas` manually; it is auto-generated.
- Do not use `GoTo` for flow control other than error handling (`On Error GoTo`).
- Do not declare site-aware variables without `As New` (e.g., `Dim x As SiteLong` is wrong; use `Dim x As New SiteLong`).
- Do not hard-code site numbers; always use `For Each site In TheExec.Sites`.
- Do not omit `ForceResults:=tlForceFlow` on `TestLimit` calls.
- Do not omit `Option Explicit` from any module.

## Templates

Generic, device-independent template files derived from production UltraFlex programs:

- [VBT Function Templates](resources/templates/vbt_templates.md) - 19 VBT function skeletons (Globals, DUTconditions class, Instruments, Exec Interpose, Continuity, Leakage, IDD, IDDQ, Scan, Logic Levels, GPIO, Trim, POR, Reference Voltage, eFuse/OTP, Utilities)
- [IG-XL Sheet Templates](resources/templates/igxl_sheets.md) - 12 DT-header sheet formats (Test Instances, Flow Table, Pin Map, Channel Map, Levels, Time Sets, Global Specs, Job List, DC/AC Specs, Load File, Sub-Program Structure, Tester Config)
- [ATP Pattern Templates](resources/templates/pattern_templates.md) - 11 ATP pattern templates (SPI Write/Read, SBC Clock, GPIO, JTAG/SWD, Keep-Alive, Frequency Counter, DSSC Scan, Flat SPI, Pipeline Flush, DUT Reset)
- [Program Review Checklist](resources/templates/program_checklist.md) - 15-section review checklist (Program Structure, Platform ID, Lifecycle, DUTconditions, Instruments, Measurement, Multi-Site, Flow, Datalogging, Patterns, Sub-Programs, Register Access, Conditional Compilation, DSP DLL, Anti-Patterns)
