"""Orchestrator for UltraFlex IG-XL multi-agent generation pipeline."""

from __future__ import annotations

import os
import ssl
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import httpx  # Required by workspace SSL policy for all new scripts.

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

from .agents import (
    ACSpecsTimingAgent,
    AgentContext,
    DatalogAnalyzerAgent,
    FlowTableAgent,
    InputSpecIngestionAgent,
    JobListAgent,
    LevelsDCSpecsAgent,
    PatternSetAgent,
    PinMapChanMapAgent,
    TestInstancesVBTAgent,
    VBTCodeWriterAgent,
    write_generated_files,
)


VALIDATION_CHECKLIST = [
    "Agent 1: Verify all pins in new device datasheet are in Pinmap",
    "Agent 1: Confirm channel assignments match HIB schematic",
    "Agent 2: Verify Vil/Vih formulas reference correct supply domain",
    "Agent 2: Confirm MDIO pin uses VT driver mode",
    "Agent 3: Confirm period selector covers interface speeds",
    "Agent 4: Validate all .PAT file paths exist on tester",
    "Agent 5: Verify continuity limits and bin assignments",
    "Agent 6: Verify OnProgramLoaded powers DIB in Exec_IP",
    "Agent 7: Verify each job references correct sheets",
    "Agent 8: Flag datalogs with >2V or <0.1V continuity",
    "Final: Load generated program in IG-XL offline mode and run Validate",
]


@dataclass
class GenerationRequest:
    device_name: str
    package_pins: int
    num_sites: int
    jobs: list[dict[str, Any]]
    base_program_path: str
    output_dir: str
    pattern_base_path: str | None = None
    pin_mapping_file: str | None = None
    datasheet_file: str | None = None
    enforce_required_vb_modules: bool = True
    required_tests: list[str] | None = None
    use_web_vbt_knowledge: bool = True


class TestProgramGeneratorOrchestrator:
    """Coordinates 8 agents to generate a new UltraFlex program bundle."""

    def __init__(self, base_program_path: str):
        self.context = AgentContext(base_program_path=Path(base_program_path))
        self.agent1 = PinMapChanMapAgent(self.context)
        self.agent2 = LevelsDCSpecsAgent(self.context)
        self.agent3 = ACSpecsTimingAgent(self.context)
        self.agent4 = PatternSetAgent(self.context)
        self.agent5 = FlowTableAgent(self.context)
        self.agent6 = TestInstancesVBTAgent(self.context)
        self.vbt_writer = VBTCodeWriterAgent(self.context)
        self.agent7 = JobListAgent(self.context)
        self.agent8 = DatalogAnalyzerAgent(self.context)
        self.input_agent = InputSpecIngestionAgent(self.context)

    @staticmethod
    def _validate_required_vb_modules(vb_bundle: dict[str, str]) -> dict[str, Any]:
        required_exact = ["Exec_IP_Module.bas", "RunVBT.bas"]
        mdio_candidates = ["MDIO_DSSC.cls", "MDIO_PA.cls", "MDIO_test.bas"]

        present = set(vb_bundle.keys())
        missing_exact = [m for m in required_exact if m not in present]
        mdio_present = any(m in present for m in mdio_candidates)

        missing = missing_exact.copy()
        if not mdio_present:
            missing.append("One of: " + ", ".join(mdio_candidates))

        return {
            "ok": len(missing) == 0,
            "missing": missing,
            "required_exact": required_exact,
            "required_mdio_any_of": mdio_candidates,
        }

    def generate(self, request: GenerationRequest) -> dict[str, Any]:
        output_dir = Path(request.output_dir)

        stats = self.agent8.extract_statistics(str(self.context.datalog_dir))
        hw_failures = self.agent8.detect_hardware_failures(str(self.context.datalog_dir))
        recommended_limits = self.agent8.recommend_limits(stats)

        pin_rows: list[dict[str, Any]] = []
        if request.pin_mapping_file:
            pin_rows = self.input_agent.parse_pin_mapping_file(
                request.pin_mapping_file,
                package_pins=request.package_pins,
                num_sites=request.num_sites,
            )

        datasheet_specs: dict[str, Any] = {}
        if request.datasheet_file:
            datasheet_specs = self.input_agent.extract_datasheet_specs(request.datasheet_file)

        pinmap = self.agent1.generate_pinmap(pin_list=pin_rows if pin_rows else None)

        pin_channel_map: dict[str, dict[str, str]] | None = None
        if pin_rows:
            pin_channel_map = {}
            for row in pin_rows:
                pin_name = row.get("pin_name", "")
                if not pin_name:
                    continue
                pin_channel_map[pin_name] = {
                    "type": row.get("type", "I/O"),
                    "package_pin": row.get("package_pin", ""),
                    "comment": row.get("comment", ""),
                }
                for site_idx in range(request.num_sites):
                    pin_channel_map[pin_name][f"site{site_idx}"] = row.get(f"site{site_idx}", "")

        chanmap = self.agent1.generate_chanmap(
            request.package_pins,
            request.num_sites,
            pin_channel_map=pin_channel_map,
        )

        package_name = f"pkg{request.package_pins}"
        vddio_domains = ["VDDIO_R", "VDDIO_M"]
        levels = self.agent2.generate_levels(package_name, vddio_domains)

        supply_voltages = {
            "VDD0P9": 0.9,
            "AVDD3P3": 3.3,
            "VDDIO_R": 3.3,
            "VDDIO_M": 3.3,
        }
        supply_voltages.update(datasheet_specs.get("supply_voltages", {}))
        dc_specs = self.agent2.generate_dc_specs(supply_voltages)

        interface_speeds = datasheet_specs.get("interface_speeds") or ["1000T", "100T", "10T"]
        ac_specs = self.agent3.generate_ac_specs(interface_speeds)

        patset = self.agent4.generate_patset(package_name, request.device_name)
        pat_missing = []
        if request.pattern_base_path:
            pat_missing = self.agent4.validate_pat_files(request.pattern_base_path)

        main_flow = self.agent5.generate_main_flow(f"{request.device_name}_A0", is_char=False)
        char_flow = self.agent5.generate_main_flow(f"{request.device_name}_A0_CHAR", is_char=True)

        test_instances = self.agent6.generate_test_instances(package_name, test_list=[])
        vb_bundle = self.agent6.generate_vbt_modules_bundle(request.device_name)
        custom_test_modules = self.vbt_writer.generate_for_required_tests(
            request.device_name,
            request.required_tests or [],
            use_web_knowledge=request.use_web_vbt_knowledge,
        )
        vb_bundle.update(custom_test_modules)
        vb_check = self._validate_required_vb_modules(vb_bundle)
        if request.enforce_required_vb_modules and not vb_check["ok"]:
            missing_text = "; ".join(vb_check["missing"])
            raise RuntimeError(
                "Required VB module check failed. Missing: " + missing_text
            )

        jobs = request.jobs or [
            {
                "name": f"{request.device_name}_A0",
                "package": package_name,
                "flow": "Flow Table",
                "is_char": False,
            },
            {
                "name": f"{request.device_name}_A0_CHAR",
                "package": package_name,
                "flow": "Flow Table Char",
                "is_char": True,
            },
        ]
        job_list = self.agent7.generate_job_list(jobs)

        generated_files = {
            "Pinmap.txt": pinmap,
            f"ChanMap_pkg{request.package_pins}_eng.txt": chanmap,
            f"Levels_pkg{request.package_pins}.txt": levels,
            "DC Specs.txt": dc_specs,
            "AC Specs.txt": ac_specs,
            f"Patset_{request.package_pins}pkg.txt": patset,
            "Flow Table.txt": main_flow,
            "Flow Table Char.txt": char_flow,
            "Test Instances Auto.txt": test_instances,
            "Job List.txt": job_list,
        }
        generated_files.update(vb_bundle)
        write_generated_files(output_dir, generated_files)

        return {
            "output_dir": str(output_dir),
            "generated_files": sorted(generated_files.keys()),
            "generated_vb_modules": sorted(vb_bundle.keys()),
            "generated_custom_test_modules": sorted(custom_test_modules.keys()),
            "vb_module_check": vb_check,
            "input_pin_rows": len(pin_rows),
            "datasheet_specs": datasheet_specs,
            "hw_failures": hw_failures,
            "recommended_limits": recommended_limits,
            "missing_patterns": pat_missing,
            "validation_checklist": VALIDATION_CHECKLIST,
        }


def build_from_reference(
    device_name: str,
    package_pins: int,
    num_sites: int,
    base_program_path: str,
    output_dir: str,
    jobs: list[dict[str, Any]] | None = None,
    pattern_base_path: str | None = None,
    pin_mapping_file: str | None = None,
    datasheet_file: str | None = None,
    enforce_required_vb_modules: bool = True,
    required_tests: list[str] | None = None,
    use_web_vbt_knowledge: bool = True,
) -> dict[str, Any]:
    """Convenience wrapper for one-shot program generation."""
    orchestrator = TestProgramGeneratorOrchestrator(base_program_path)
    request = GenerationRequest(
        device_name=device_name,
        package_pins=package_pins,
        num_sites=num_sites,
        jobs=jobs or [],
        base_program_path=base_program_path,
        output_dir=output_dir,
        pattern_base_path=pattern_base_path,
        pin_mapping_file=pin_mapping_file,
        datasheet_file=datasheet_file,
        enforce_required_vb_modules=enforce_required_vb_modules,
        required_tests=required_tests,
        use_web_vbt_knowledge=use_web_vbt_knowledge,
    )
    return orchestrator.generate(request)
