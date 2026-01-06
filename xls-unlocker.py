#!/usr/bin/env python3
"""
Legacy Excel XLS Unlocker / DPB Record Modifier
"""
import argparse
import shutil
from pathlib import Path

TARGET = b"DPB"
REPLACEMENT = b"DPx"

def modify_xls(input_path: Path, output_path: Path) -> None:
    data = input_path.read_bytes()
    occurrences = []
    start = 0
    while True:
        idx = data.find(TARGET, start)
        if idx == -1:
            break
        occurrences.append(idx)
        start = idx + 1

    if not occurrences:
        raise RuntimeError("No DPB records found. File may not be legacy XLS.")

    print(f"Found {len(occurrences)} DPB record(s):")
    for offset in occurrences:
        print(f"  Offset 0x{offset:08X}")

    modified = bytearray(data)
    for offset in occurrences:
        modified[offset : offset + len(TARGET)] = REPLACEMENT

    output_path.write_bytes(modified)
    print(f"Modified file written to: {output_path}")

def main():
    parser = argparse.ArgumentParser(
        description="Modify DPB records in legacy Excel .xls files"
    )
    parser.add_argument("xls", type=Path, help="Input .xls file")
    parser.add_argument(
        "--out",
        type=Path,
        help="Output file path (default: <input>.patched.xls)",
    )
    args = parser.parse_args()

    if args.xls.suffix.lower() != ".xls":
        raise ValueError("This tool supports legacy .xls files only.")

    output = args.out or args.xls.with_name(args.xls.stem + ".patched.xls")
    shutil.copyfile(args.xls, output)
    modify_xls(output, output)

if __name__ == "__main__":
    main()
