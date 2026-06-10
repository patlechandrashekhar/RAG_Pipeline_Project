# Test Program MCP Skills (Starter)

This file is a starter template for skill behavior when using the `ultraflex-testprog` MCP.

## Primary workflow

1. Ask user for channel assignment input.
2. Call `build_pinmap_package` with:
   - `channel_assignment`
   - `project_name`
   - `default_instrument`
3. When user asks for IG-XL seed artifacts, call `generate_igxl_seed_package` with:
   - `channel_assignment`
   - `project_name`
   - `default_instrument`
   - `default_test_name`
   - `default_levelset_name`
   - `default_timingset_name`
   - `default_period_ns`
   - `protocol_preset` (SPI/I2C/UART/JTAG)
3. Present:
   - Pin map
   - Channel map
   - Datasheet draft
4. If warnings are returned, show warnings first and request correction.

## Acting rules for LLM

- Do not invent channel IDs, pin names, or instrument limits.
- Preserve user-provided naming exactly unless user asks to normalize.
- When fields are missing, state assumptions explicitly.
- Show conflicts (duplicate channel usage, duplicate pin rows) before generating final files.
- Keep generated data in plain CSV/Markdown for easy import into IG-XL workflows.
- If `protocol_preset` is provided, prefer preset timing/level defaults over blank placeholders.

## Input flexibility (temporary)

Until a strict input schema is finalized, accept:

- JSON list/object assignments
- CSV with headers
- Simple token lines like: `PIN=RESET_N CH=CH01 SITE=1 DIR=IN`

## Planned extension points

- Standard JSON schema validation (future)
- Pin-group to level-set/timing-set mapping
- Test-instance skeleton generation from pin groups
- Limits-table seed generation
- Per-interface Timing/Levels auto values (based on protocol templates)
