#!/usr/bin/env python3
"""Check JSON register names/offsets against the RTL AXI read muxes."""
import argparse
import json
import re
from pathlib import Path


def rtl_registers(path):
    text = Path(path).read_text(encoding="utf-8", errors="replace")
    registers = {}
    for index, expression in re.findall(
        r"'d(\d+)\s*:\s*reg_read_data\s*=\s*([^;]+);", text
    ):
        match = re.fullmatch(r"\s*([A-Za-z_]\w*)_d\s*", expression)
        registers[int(index) * 4] = match.group(1) if match else "__DIRECT__"
    return registers


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("schema", type=Path)
    parser.add_argument("interface_rtl", type=Path)
    parser.add_argument("accumulator_rtl", type=Path)
    args = parser.parse_args()
    schema = json.loads(args.schema.read_text(encoding="utf-8"))
    blocks = {b["name"]: b for b in schema["blocks"]}
    expected_if = {}
    for name in ("IO_DELAY_CTRL", "DMA", "SPI", "PIN_CTRL"):
        for reg in blocks[name]["registers"]:
            expected_if[int(reg["offset"], 0)] = reg["name"]
    expected_acc = {
        int(reg["offset"], 0): reg["name"]
        for reg in blocks["ACCUMULATOR"]["registers"]
    }
    errors = []
    for label, expected, actual in (
        ("interface", expected_if, rtl_registers(args.interface_rtl)),
        ("accumulator", expected_acc, rtl_registers(args.accumulator_rtl)),
    ):
        if expected != actual:
            for offset in sorted(set(expected) | set(actual)):
                actual_name = actual.get(offset)
                if expected.get(offset) != actual_name and actual_name != "__DIRECT__":
                    errors.append(
                        f"{label} 0x{offset:04X}: JSON={expected.get(offset)} RTL={actual.get(offset)}"
                    )
    if errors:
        raise SystemExit("\n".join(errors))
    print(f"REGMAP_RTL_MATCH interface={len(expected_if)} accumulator={len(expected_acc)}")


if __name__ == "__main__":
    main()
