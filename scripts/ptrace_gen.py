#!/usr/bin/env python3
"""Generate a HotSpot power trace from an Excel phase table.

The input workbook contains one ``prefill`` row and one
``decode_per_token`` row. Each phase is sampled at 0.001 seconds by default.
Prefill runs once. Decode duration is 129 times the per-token latency;
each stage repeats its power row ceil(stage_duration / sampling_interval) times.

* ``<output>.ptrace`` contains the HotSpot header and power values only.
* ``<output>.time.tsv`` contains the phase duration and cumulative time for
  every sample.
* ``<output>.units.tsv`` maps each ptrace column to its LCF layer and FLP unit.

The LCF, rather than a fixed template, defines the power-column layout. Every
layer whose power flag is ``Y`` contributes the units from its referenced FLP,
in LCF order. Repeated references to an FLP remain separate column groups.

Phase powers are calculated from Excel energy (joules) divided by latency
(seconds). Flash, PB, and Other totals are shared across all matching units
in all power-enabled layers. Unmapped units receive zero power.
"""

from __future__ import annotations

import argparse
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import openpyxl
except ImportError as exc:  # pragma: no cover - depends on the caller's environment
    raise SystemExit(
        "error: reading .xlsx input requires openpyxl; install it with "
        "'python -m pip install openpyxl'"
    ) from exc


DEFAULT_DECODE_COUNT = 129
DEFAULT_SAMPLING_INTERVAL = 0.001


@dataclass(frozen=True)
class PhaseRecord:
    """One phase row read from the Excel input."""

    phase: str
    latency_s: float
    source_values: Dict[str, Any]


@dataclass(frozen=True)
class PowerUnit:
    """One power column in LCF layer and FLP unit order."""

    layer_number: int
    flp_path: Path
    flp_reference: str
    unit_name: str


def _normalise_header(value: Any) -> str:
    return str(value).strip().casefold() if value is not None else ""


def _as_float(value: Any, label: str) -> float:
    if value is None or isinstance(value, bool):
        raise ValueError(f"{label} must be a number, got {value!r}")
    try:
        number = float(str(value).strip().rstrip("sS"))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a number, got {value!r}") from exc
    if not math.isfinite(number):
        raise ValueError(f"{label} must be finite, got {number}")
    if number < 0:
        raise ValueError(f"{label} must be non-negative, got {number}")
    return number


def _canonical_phase(value: Any, row_number: int) -> str:
    phase = str(value).strip().casefold() if value is not None else ""
    if phase == "prefill":
        return "prefill"
    if phase in {"decode", "decode_per_token", "decode per token"}:
        return "decode_per_token"
    raise ValueError(
        f"row {row_number}: phase must be 'prefill' or 'decode_per_token', got {value!r}"
    )


def _first_nonempty_row(rows: Iterable[Sequence[Any]]) -> Tuple[int, List[Any]]:
    for row_number, row in enumerate(rows, start=1):
        values = list(row)
        if any(value is not None and str(value).strip() for value in values):
            return row_number, values
    raise ValueError("Excel input does not contain a header row")


def read_excel_records(path: Path, sheet_name: Optional[str] = None) -> Dict[str, PhaseRecord]:
    """Read the single prefill and decode rows from an Excel workbook."""

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if sheet_name is None:
            worksheet = workbook.active
        elif sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
        else:
            available = ", ".join(workbook.sheetnames)
            raise ValueError(f"worksheet {sheet_name!r} not found; available sheets: {available}")

        rows = worksheet.iter_rows(values_only=True)
        header_row_number, header_values = _first_nonempty_row(rows)
        headers = [str(value).strip() if value is not None else "" for value in header_values]
        normalised = [_normalise_header(value) for value in headers]

        try:
            phase_index = normalised.index("phase")
        except ValueError as exc:
            raise ValueError("Excel input must contain a 'phase' column") from exc

        latency_candidates = {"latency", "latency_s", "time", "time_s"}
        latency_index = next(
            (index for index, value in enumerate(normalised) if value in latency_candidates),
            None,
        )
        if latency_index is None:
            raise ValueError("Excel input must contain a 'latency' or 'latency_s' column")

        records: Dict[str, PhaseRecord] = {}
        for row_number, row in enumerate(rows, start=header_row_number + 1):
            values = list(row)
            if not any(value is not None and str(value).strip() for value in values):
                continue

            phase_value = values[phase_index] if phase_index < len(values) else None
            if phase_value is None or not str(phase_value).strip():
                continue
            phase = _canonical_phase(phase_value, row_number)
            if phase in records:
                raise ValueError(f"Excel input contains more than one {phase!r} row")

            latency_value = values[latency_index] if latency_index < len(values) else None
            latency_s = _as_float(latency_value, f"row {row_number} latency")
            source_values = {
                header: (values[index] if index < len(values) else None)
                for index, header in enumerate(headers)
                if header
            }
            records[phase] = PhaseRecord(phase, latency_s, source_values)
    finally:
        workbook.close()

    missing = [phase for phase in ("prefill", "decode_per_token") if phase not in records]
    if missing:
        raise ValueError("Excel input is missing phase row(s): " + ", ".join(missing))
    return records


def read_flp_units(path: Path) -> List[str]:
    """Return floorplan unit names in the order declared by the FLP."""

    units: List[str] = []
    with path.open(encoding="utf-8") as stream:
        for raw_line in stream:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            units.append(line.split()[0])
    if not units:
        raise ValueError(f"floorplan has no units: {path}")
    return units


def read_lcf_power_units(lcf_path: Path) -> List[PowerUnit]:
    """Expand every power-enabled LCF layer into ordered ptrace columns."""

    with lcf_path.open(encoding="utf-8") as stream:
        fields = [
            (line_number, raw_line.split("#", 1)[0].strip())
            for line_number, raw_line in enumerate(stream, start=1)
            if raw_line.split("#", 1)[0].strip()
        ]

    records = []
    field_index = 0
    while field_index < len(fields):
        remaining = len(fields) - field_index
        if remaining < 6:
            line_number = fields[field_index][0]
            raise ValueError(f"{lcf_path}:{line_number}: incomplete LCF layer record")

        # HotSpot accepts either numeric heat capacity + resistivity + thickness
        # (seven fields) or a material name + thickness (six fields).
        try:
            float(fields[field_index + 3][1])
            field_count = 7
        except ValueError:
            field_count = 6
        if remaining < field_count:
            line_number = fields[field_index][0]
            raise ValueError(f"{lcf_path}:{line_number}: incomplete LCF layer record")

        records.append(fields[field_index : field_index + field_count])
        field_index += field_count

    layer_count = len(records)
    layer_numbers = []
    for record in records:
        layer_line = record[0][0]
        try:
            layer_number = int(record[0][1])
        except ValueError as exc:
            raise ValueError(
                f"{lcf_path}:{layer_line}: invalid layer number {record[0][1]!r}"
            ) from exc
        if layer_number < 0:
            raise ValueError(f"{lcf_path}:{layer_line}: layer number must be non-negative")
        layer_numbers.append(layer_number)

    duplicates = sorted({n for n in layer_numbers if layer_numbers.count(n) > 1})
    missing = sorted(set(range(layer_count)) - set(layer_numbers))
    if duplicates or missing or layer_numbers != list(range(layer_count)):
        details = []
        if duplicates:
            details.append(f"duplicate layer number(s): {duplicates}")
        if missing:
            details.append(f"missing layer number(s): {missing}")
        if not duplicates and not missing:
            details.append("layer records are not in ascending layer-number order")
        raise ValueError(
            f"{lcf_path}: layer numbers must be unique, contiguous, and listed in order; "
            + "; ".join(details)
        )

    power_units: List[PowerUnit] = []
    for record_index, record in enumerate(records):
        layer_number = layer_numbers[record_index]
        power_flag = record[2][1].strip().upper()
        if power_flag not in {"Y", "N"}:
            raise ValueError(
                f"{lcf_path}:{record[2][0]}: power flag must be Y or N, got {power_flag!r}"
            )
        if power_flag != "Y":
            continue

        flp_reference = record[-1][1]
        flp_path = Path(flp_reference)
        if not flp_path.is_absolute():
            flp_path = lcf_path.parent / flp_path
        flp_path = flp_path.resolve()
        if not flp_path.is_file():
            raise ValueError(
                f"{lcf_path}:{record[-1][0]}: floorplan for power layer {layer_number} "
                f"does not exist: {flp_path}"
            )

        for unit_name in read_flp_units(flp_path):
            power_units.append(
                PowerUnit(
                    layer_number=layer_number,
                    flp_path=flp_path,
                    flp_reference=flp_reference,
                    unit_name=unit_name,
                )
            )

    if not power_units:
        raise ValueError(f"LCF contains no layers with power dissipation enabled: {lcf_path}")
    return power_units


def expand_records(
    records: Dict[str, PhaseRecord], decode_count: int = DEFAULT_DECODE_COUNT,
    sampling_interval: float = DEFAULT_SAMPLING_INTERVAL,
) -> List[PhaseRecord]:
    """Repeat each stage's power for ceil(total stage duration / interval) steps.

    Power continues to use the original Excel latency; rounding the duration
    up may add less than one sampling interval to each stage.
    """

    if decode_count < 0:
        raise ValueError(f"decode_count must be non-negative, got {decode_count}")
    if not math.isfinite(sampling_interval) or sampling_interval <= 0:
        raise ValueError("sampling_interval must be finite and positive")
    steps = {}
    for phase in ("prefill", "decode_per_token"):
        latency = records[phase].latency_s
        if not math.isfinite(latency) or latency <= 0:
            raise ValueError(f"{phase}: latency must be finite and positive")
        steps[phase] = math.ceil(latency / sampling_interval)
    decode_steps = math.ceil(records["decode_per_token"].latency_s * decode_count / sampling_interval)
    return ([records["prefill"]] * steps["prefill"]
            + [records["decode_per_token"]] * decode_steps)


def build_power_row(power_units: Sequence[PowerUnit], record: PhaseRecord) -> List[float]:
    """Apply the component energy mapping across all active LCF layers.

    NPU components receive their specified E/latency individually. Flash
    blocks, Logic PB blocks, and Logic Other blocks share their respective
    total across every matching column, including repeated FLP layers.
    """

    latency = _as_float(record.latency_s, f"{record.phase} latency")
    if latency <= 0:
        raise ValueError(f"{record.phase}: latency must be positive")

    fields = {_normalise_header(key): value for key, value in record.source_values.items()}

    def phase_power(energy_name: str) -> float:
        # The supplied workbook uses E_*_J; also accept E_* headers.
        for key in (energy_name.casefold() + "_j", energy_name.casefold()):
            if key in fields:
                energy = _as_float(fields[key], f"{record.phase} {energy_name}")
                return energy / latency
        raise ValueError(f"{record.phase}: missing Excel energy column {energy_name}_J")

    groups: Dict[str, List[int]] = {"flash": [], "pb": [], "other": []}
    npu_indices: Dict[str, List[int]] = {"pu_buf": [], "mmu": [], "npu_ctrl": []}
    for index, unit in enumerate(power_units):
        floorplan = unit.flp_path.stem.casefold()
        name = unit.unit_name.casefold().replace("-", "_")
        if floorplan == "flash" and name.startswith("flash_") and name != "flash_tsv":
            groups["flash"].append(index)
        elif floorplan == "logic":
            if name.startswith("pb_") or name == "pb":
                groups["pb"].append(index)
            elif name.startswith("other_") or name == "other":
                groups["other"].append(index)
        elif floorplan == "npu" and name in npu_indices:
            npu_indices[name].append(index)

    powers = [0.0] * len(power_units)
    for group, energy_name in (("flash", "E_FA"), ("pb", "E_WB"), ("other", "E_ECC")):
        indices = groups[group]
        if indices:
            shared_power = phase_power(energy_name) / len(indices)
            for index in indices:
                powers[index] = shared_power

    for component, energy_name in (
        ("pu_buf", "E_WB"), ("mmu", "E_Flash_MAC"), ("npu_ctrl", "E_Ctrl_iface")
    ):
        indices = npu_indices[component]
        if indices:
            component_power = phase_power(energy_name)
            for index in indices:
                powers[index] = component_power
    return powers


def write_ptrace(
    path: Path,
    power_units: Sequence[PowerUnit],
    trace_records: Sequence[PhaseRecord],
) -> None:
    """Write a standard HotSpot ptrace with one power row per trace record."""

    unit_names = [unit.unit_name for unit in power_units]
    header_line = "\t".join(unit_names)
    # Calculate each phase once and stream repeated text rows to keep memory
    # bounded when short sampling intervals produce many thousands of rows.
    power_rows: Dict[str, str] = {}
    for record in trace_records:
        if record.phase in power_rows:
            continue
        powers = build_power_row(power_units, record)
        if len(powers) != len(power_units):
            raise ValueError(
                f"power row has {len(powers)} values, expected {len(power_units)}"
            )
        if any(not math.isfinite(float(value)) or float(value) < 0 for value in powers):
            raise ValueError("power values must be finite and non-negative")
        power_rows[record.phase] = "\t".join(f"{value:.12g}" for value in powers) + "\n"

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write(header_line + "\n")
        for record in trace_records:
            stream.write(power_rows[record.phase])


def write_time_trace(
    path: Path, trace_records: Sequence[PhaseRecord],
    sampling_interval: float = DEFAULT_SAMPLING_INTERVAL,
) -> None:
    """Write actual fixed-step intervals matching ptrace rows."""

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write("trace_index\tphase\tduration_s\tstart_time_s\tend_time_s\n")
        for index, record in enumerate(trace_records):
            start = index * sampling_interval
            elapsed = (index + 1) * sampling_interval
            stream.write(
                f"{index}\t{record.phase}\t{sampling_interval:.12g}\t"
                f"{start:.12g}\t{elapsed:.12g}\n"
            )


def write_phase_ptrace(
    path: Path, power_units: Sequence[PowerUnit], record: PhaseRecord,
    rows: int,
) -> None:
    """Write one phase-only ptrace with a repeated constant power row."""
    if rows <= 0:
        raise ValueError(f"{record.phase}: phase must contain at least one sample")
    unit_names = [unit.unit_name for unit in power_units]
    powers = build_power_row(power_units, record)
    if len(powers) != len(power_units):
        raise ValueError(f"power row has {len(powers)} values, expected {len(power_units)}")
    if any(not math.isfinite(float(value)) or float(value) < 0 for value in powers):
        raise ValueError("power values must be finite and non-negative")
    line = "\t".join(f"{value:.12g}" for value in powers) + "\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write("\t".join(unit_names) + "\n")
        stream.write(line * rows)


def write_units_manifest(path: Path, power_units: Sequence[PowerUnit]) -> None:
    """Write LCF layer and FLP identity for every ptrace power column."""

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write("column_index\tlayer_number\tflp_reference\tunit_name\n")
        for index, unit in enumerate(power_units):
            stream.write(
                f"{index}\t{unit.layer_number}\t{unit.flp_reference}\t{unit.unit_name}\n"
            )


def generate_from_excel(
    input_xlsx: Path,
    lcf_path: Path,
    output_ptrace: Path,
    output_time: Path,
    output_units: Path,
    output_prefill: Optional[Path] = None,
    output_decode: Optional[Path] = None,
    sheet_name: Optional[str] = None,
    decode_count: int = DEFAULT_DECODE_COUNT,
    sampling_interval: float = DEFAULT_SAMPLING_INTERVAL,
) -> Tuple[int, int]:
    """Read Excel input and write the ptrace, time, and unit-map files."""

    records = read_excel_records(input_xlsx, sheet_name)
    power_units = read_lcf_power_units(lcf_path)
    prefill_steps = math.ceil(records["prefill"].latency_s / sampling_interval)
    decode_steps = math.ceil(records["decode_per_token"].latency_s * decode_count / sampling_interval)
    trace_records = expand_records(records, decode_count, sampling_interval)
    if output_prefill is not None and output_decode is not None:
        write_phase_ptrace(output_prefill, power_units, records["prefill"], prefill_steps)
        write_phase_ptrace(output_decode, power_units, records["decode_per_token"], decode_steps)
    else:
        write_ptrace(output_ptrace, power_units, trace_records)
    prefill_records = trace_records[:prefill_steps]
    decode_records = trace_records[prefill_steps:]
    write_time_trace(
        output_time.with_name(output_time.stem + "_prefill.tsv"),
        prefill_records, sampling_interval,
    )
    write_time_trace(
        output_time.with_name(output_time.stem + "_decode.tsv"),
        decode_records, sampling_interval,
    )
    write_units_manifest(output_units, power_units)
    return len(power_units), len(trace_records)


def main(argv: Optional[Sequence[str]] = None) -> int:
    repo_root = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(
        description="Generate a HotSpot ptrace and matching time files from Excel and LCF inputs."
    )
    parser.add_argument(
        "--input-xlsx",
        "--input",
        dest="input_xlsx",
        required=True,
        type=Path,
        help="Excel workbook containing prefill and decode_per_token rows",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="worksheet name (default: active worksheet)",
    )
    parser.add_argument(
        "--lcf",
        "--grid-layer-file",
        dest="lcf",
        required=True,
        type=Path,
        help="HotSpot LCF; determines active power units and their order",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=repo_root / "ptrace",
        help="directory for default output files",
    )
    parser.add_argument(
        "--output-ptrace",
        type=Path,
        default=None,
        help="output ptrace path (default: <output-dir>/<lcf-directory-name>.ptrace)",
    )
    parser.add_argument(
        "--output-prefill", type=Path, default=None,
        help="prefill-only ptrace path (default: <output-dir>/<lcf-directory-name>_prefill.ptrace)",
    )
    parser.add_argument(
        "--output-decode", type=Path, default=None,
        help="decode-only ptrace path (default: <output-dir>/<lcf-directory-name>_decode.ptrace)",
    )
    parser.add_argument(
        "--output-time",
        type=Path,
        default=None,
        help="output time path (default: <output-dir>/<lcf-directory-name>.time.tsv)",
    )
    parser.add_argument(
        "--output-units",
        type=Path,
        default=None,
        help="output unit map (default: <output-dir>/<lcf-directory-name>.units.tsv)",
    )
    parser.add_argument(
        "--decode-count",
        type=int,
        default=DEFAULT_DECODE_COUNT,
        help=f"number of decode tokens (default: {DEFAULT_DECODE_COUNT})",
    )
    parser.add_argument(
        "--sampling-intvl", type=float, default=DEFAULT_SAMPLING_INTERVAL,
        help="seconds per power row; must match HotSpot config (default: 0.001)",
    )
    args = parser.parse_args(argv)

    if not args.input_xlsx.is_file():
        parser.error(f"input Excel file not found: {args.input_xlsx}")
    if not args.lcf.is_file():
        parser.error(f"LCF file not found: {args.lcf}")

    stem = args.lcf.resolve().parent.name
    output_ptrace = args.output_ptrace or args.output_dir / f"{stem}.ptrace"
    output_prefill = args.output_prefill or args.output_dir / f"{stem}_prefill.ptrace"
    output_decode = args.output_decode or args.output_dir / f"{stem}_decode.ptrace"
    output_time = args.output_time or args.output_dir / f"{stem}.time.tsv"
    output_units = args.output_units or args.output_dir / f"{stem}.units.tsv"

    try:
        unit_count, trace_count = generate_from_excel(
            input_xlsx=args.input_xlsx,
            lcf_path=args.lcf,
            output_ptrace=output_ptrace,
            output_time=output_time,
            output_units=output_units,
            output_prefill=output_prefill,
            output_decode=output_decode,
            sheet_name=args.sheet,
            decode_count=args.decode_count,
            sampling_interval=args.sampling_intvl,
        )
    except (OSError, ValueError, NotImplementedError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    print(f"wrote {output_prefill} (prefill ptrace, {unit_count} units)")
    print(f"wrote {output_decode} (decode ptrace, {unit_count} units)")
    print(f"wrote {output_time.with_name(output_time.stem + '_prefill.tsv')}")
    print(f"wrote {output_time.with_name(output_time.stem + '_decode.tsv')}")
    print(f"wrote {output_units} ({unit_count} LCF-to-column mappings)")
    print(f"sampling interval: {args.sampling_intvl:g} s (must match HotSpot config)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
