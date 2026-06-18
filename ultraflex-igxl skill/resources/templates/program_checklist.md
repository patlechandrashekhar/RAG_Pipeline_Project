# UltraFlex/VBT Program Review Checklist

Use this checklist to audit UltraFlex/IG-XL test programs for
correctness, completeness, and best practices.

## 1. Program Structure

- [ ] Every VBA module has `Option Explicit` as the first line
- [ ] Every top-level test function uses `TF_` prefix and returns `Long`
- [ ] Every VBT function has `On Error GoTo errHandler` as first executable line
- [ ] Every error handler uses pattern: `TheExec.Datalog.WriteComment "Error: " & Err.Number & " " & Err.Description` followed by `If AbortTest Then Exit Function Else Resume Next`
- [ ] A `Globals` module exists with `TestCondition` type and test condition arrays
- [ ] A `DUTconditions` class exists with EOS-safe supply ramping
- [ ] An `Instruments` module exists with range enums and init functions
- [ ] An `Exec_IP_Module` exists with all lifecycle hooks
- [ ] A `Utilities` module exists with shared helper functions
- [ ] `RunVBT.bas` is not manually edited (auto-generated wrappers)
- [ ] Owner attribution comments present: `'''Owner: <name>`
- [ ] Variable naming follows convention: `g_` globals, `c_` constants, `m_` private members
- [ ] Site-aware variables declared with `As New` (`Dim x As New SiteDouble`)
- [ ] No scalar `Double`/`Long` used where `SiteDouble`/`SiteLong` is required
- [ ] Timer profiling present: `theTimer.VBTFunctionStart` / `theTimer.VBTFunctionFinish`

## 2. UltraFlex Platform Identification

- [ ] Workbook sheets use `DT<SheetType>` headers with `platform=Jaguar`
- [ ] Not using MicroFlex `DFF` headers
- [ ] Sheet metadata includes correct version numbers

## 3. Exec Interpose Lifecycle

- [ ] `OnProgramLoaded`: Sets program working directory, adds register map library reference (if using vblite)
- [ ] `OnProgramValidated`: Sets `g_GlobalVarReset = True`, enables/disables sites
- [ ] `OnProgramValidated`: Calls `tl_PinListDataSort(False)`
- [ ] `OnProgramValidated`: Configures TDR exclusion pins
- [ ] `OnProgramStarted`: Checks `g_GlobalVarReset` flag and reinits if needed
- [ ] `OnProgramStarted`: Initializes timer, pattern directories, register map objects (if using vblite)
- [ ] `OnProgramStarted`: Powers on DIB in correct sequence (5V relays before 15V)
- [ ] `OnProgramStarted`: Configures datalog setup
- [ ] `OnProgramEnded`: Powers down DUT safely, opens all relays
- [ ] `OnProgramEnded`: Disconnects differential meter
- [ ] `OnGlobalVariableReset`: Sets `g_GlobalVarReset = False`, terminates COM objects
- [ ] `OnAlarmOccurred`: Parses and logs alarm details
- [ ] `OnPreShutDownSite`: Powers down site-specific resources
- [ ] `OnPostShutDownSite`: Calls `theTimer.SiteShutdown`

## 4. DUTconditions Class

- [ ] EOS-safe supply ramping implemented (exponential step algorithm)
- [ ] Absolute maximum rating checks before setting supply voltage
- [ ] VIH/VIL safety checks (VIH-VIL gap maintained, neither exceeds supply)
- [ ] `saveState()` / `restoreState()` properly preserve all supply and level values
- [ ] `powerOn()` ramps supplies in correct order (DVDD before AVDD before IOVDD)
- [ ] `powerOff()` sets supplies to 0V in reverse order
- [ ] `SetSettlingTimer` called after supply changes
- [ ] Reset pin held high via PPMU during power-up sequence

## 5. Instrument Usage

### DCVI (DC-30/DC-75/UltraVI80)

- [ ] Current range set explicitly before measurement
- [ ] Voltage range set explicitly when measuring voltages
- [ ] Meter mode set correctly (`tlDCVIMeterCurrent` or `tlDCVIMeterVoltage`)
- [ ] Meter filter configured appropriately
- [ ] Compliance range set for supply pins
- [ ] Connect/Gate in correct order (Connect before Gate = True)
- [ ] Gate cleanup at function exit
- [ ] Bandwidth reduced for low-current measurements
- [ ] UltraVI80: BleederResistor configured as needed
- [ ] UltraVI80: HiZ gate mode used where appropriate (`tlDCVIGateOffHiZ`)

### PPMU (UltraPin800)

- [ ] PPMU clamps set appropriately (`SetClampsVHi`, `SetClampsVLo`)
- [ ] Current range appropriate (2uA-50mA on UltraPin800)
- [ ] Gate set to `tlOn` after Connect
- [ ] Disconnect called on cleanup
- [ ] `ForceV 0` applied between measurements to prevent EOS

### Digital (UltraPin800/GigaDig)

- [ ] `ApplyLevelsTiming` called with correct parameters after level changes
- [ ] `DisconnectPins` called before PPMU/DCVI measurements on digital pins
- [ ] `ConnectPins` called after measurement cleanup
- [ ] `PatternBurstPassed(site)` or `FailCountEx(site)` checked per site after pattern runs
- [ ] `Patgen.HaltWait` called after `Patterns.Pat().Start()`
- [ ] HRAM configured appropriately for debug patterns
- [ ] `ForcedHVPoweredMode` set only when needed and documented
- [ ] Pipeline flush (`pipe_minus`) used in patterns where immediate timing is critical

### PLMeter / Differential Meter

- [ ] PLMeter mode set correctly (Direct, Differential, Precision)
- [ ] Sample rate and size configured
- [ ] Filter and filter delay set
- [ ] `AlarmLatching = True` for robust measurement
- [ ] Differential meter offset calibrated during board checker

### Relays

- [ ] Settling delay after every relay state change
- [ ] All relays opened in `OnProgramEnded`
- [ ] DIB power sequence: 5V relay supplies before 15V relay supplies

## 6. Measurement Practices

- [ ] PLC-synchronized sample count for supply current measurements
- [ ] Adequate settling time after supply or current range changes
- [ ] Checkerboard pin grouping for continuity and leakage (grp_OS_1, grp_OS_2)
- [ ] Adjacent pins forced to different voltages during leakage testing
- [ ] Offline mode guard with simulated values
- [ ] `FailCountEx(site)` used for per-site per-pin failure analysis

## 7. Multi-Site

- [ ] Site loops use `For Each site In TheExec.Sites`
- [ ] No hardcoded site numbers
- [ ] `SiteDouble`/`SiteLong`/`SiteBoolean` used consistently
- [ ] `MSV()` function used for site-uniform initialization (`Set x = MSV(0)`)
- [ ] Site selection saved/restored when temporarily modified
- [ ] Deferred math operations evaluated before results are needed (avoid memory leaks)
- [ ] Global SiteVariants set to `Nothing` in `OnProgramEnded` to prevent memory leaks

## 8. Flow Table

- [ ] Calibration phase gated by `Calibrate` enable word
- [ ] Production phase gated by `Calok` enable word
- [ ] `set-error-bin` opcode at start of each phase
- [ ] Every `Test` opcode followed by `Use-Limit` rows
- [ ] Calibration failures bin to dedicated error bin (e.g., 31/99)
- [ ] Production failures use categorized bins
- [ ] Test numbers sequential without gaps
- [ ] Concurrent subflows properly defined with `concurrent` / `concurrent-end` opcodes
- [ ] Sub-flow calls use `call` opcode correctly
- [ ] Limits use appropriate units and scale factors

## 9. Datalogging

- [ ] `ForceResults:=tlForceFlow` on all `TestLimit` calls
- [ ] Scale types match measurements (`scaleMicro` for uA, `scaleNano` for nA)
- [ ] `DisablePinNameInPTR` toggled for PinListData results
- [ ] Dynamic test names set with `TName:=` parameter where flow table names are insufficient
- [ ] `Formatstr` parameter used for consistent result formatting
- [ ] Deferred evaluation results (background DSP) evaluated before `TestLimit`

## 10. Pattern Files

- [ ] Every `.pat` has a companion `.pvs` file
- [ ] `digital_inst = hsd` present in ATP headers (UltraFlex-specific)
- [ ] `opcode_mode` matches timing mode (single/dual/dual_2x)
- [ ] SBC drive format used for clock pins where appropriate
- [ ] `halt` terminates every pattern sequence
- [ ] Pin names match Pin Map exactly
- [ ] Time set names match Time Sets sheet
- [ ] `pipe_minus` used where immediate timing is critical
- [ ] `stv` markers placed correctly for DSSC capture
- [ ] Subroutine imports match available `.pat` subroutine files

## 11. Sub-Program Organization (UltraFlex-specific)

- [ ] Sub-programs in `sp_*` directories with `.igxlsp` project files
- [ ] Each sub-program has its own test instances, flows, and pattern directories
- [ ] Sub-program VBT modules in `sp_*/src/` directory
- [ ] Sub-program flows callable from main flow via `call` opcode
- [ ] No circular dependencies between sub-programs
- [ ] Sub-program test instance sheets listed in job list

## 12. Register Access

- [ ] Register communication method established (vblite-generated APIs, pattern-based SPI/JTAG, or custom VBA class)
- [ ] Register read/write functions used consistently throughout program
- [ ] Key unlock sequences executed before protected register access
- [ ] Register verification flag used appropriately (NOVERIF vs VERIFY)
- [ ] JTAG/SWD communication class properly initialized and terminated

## 13. Conditional Compilation

- [ ] Conditional compile arguments defined in Project Properties
- [ ] `#If (PGM_MASK And FT) Then` pattern used for probe vs. FT code
- [ ] Debug output gated by `#If (SYS_DEBUG = 1) Then`
- [ ] No probe-specific code executing during final test and vice versa

## 14. C++ Custom DSP DLL (if applicable)

- [ ] Functions exported with `TL_DSP_LIB` macro
- [ ] First parameter is `long *error_code`
- [ ] VARIANT types validated before use
- [ ] Memory freed explicitly
- [ ] Version info exposed via `get_version()` function
- [ ] VBA wrapper class uses `LoadLibrary`/`FreeLibrary`

## 15. Common Anti-Patterns

- [ ] No hardcoded limits in VBT code (use `tlForceFlow`)
- [ ] No global scalar variables where site-aware types should be used
- [ ] No `Resume Next` without proper error handler
- [ ] No missing cleanup paths
- [ ] No relay toggling without settling delays
- [ ] No voltage forcing on disconnected pins
- [ ] No pattern runs without checking pass/fail afterward
- [ ] No deferred SiteVariant math assigned to global variables without evaluation (memory leak)
- [ ] No manual edits to `RunVBT.bas`
- [ ] No VBT function names resembling Excel cell references (e.g., `A1`, `AB12`)
