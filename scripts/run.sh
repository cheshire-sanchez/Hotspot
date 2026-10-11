#!/usr/bin/env bash
set -euo pipefail

# Usage: bash scripts/run.sh [design-directory] [Excel-file-in-psource]
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DESIGN="${1:-$REPO_ROOT/HBFuse-4}"
DESIGN="$(cd "$DESIGN" && pwd)"
DESIGN_NAME="$(basename "$DESIGN")"
SOURCE_DIR="$REPO_ROOT/psource"
PTRACE_DIR="$REPO_ROOT/ptrace"
OUT_BASE="$REPO_ROOT/outputs/$DESIGN_NAME"
HOTSPOT="${HOTSPOT:-$REPO_ROOT/hotspot}"
SAMPLING_INTVL=0.001
DECODE_COUNT=129

if command -v python3 >/dev/null 2>&1; then
  PYTHON=python3
elif command -v python >/dev/null 2>&1; then
  PYTHON=python
else
  echo "error: Python not found" >&2
  exit 1
fi

EXCEL="$SOURCE_DIR/$(basename "${2:-W4_power_4.xlsx}")"
for file in "$EXCEL" "$DESIGN/example.lcf" "$DESIGN/example.config" "$DESIGN/example.materials" "$HOTSPOT"; do
  if [[ ! -f "$file" ]]; then
    echo "error: file not found: $file" >&2
    exit 1
  fi
done

mkdir -p "$PTRACE_DIR" "$OUT_BASE"
PREFILL_TRACE="$PTRACE_DIR/${DESIGN_NAME}_prefill.ptrace"
DECODE_TRACE="$PTRACE_DIR/${DESIGN_NAME}_decode.ptrace"
TIME_PREFIX="$PTRACE_DIR/${DESIGN_NAME}.time.tsv"
UNITS_FILE="$PTRACE_DIR/${DESIGN_NAME}.units.tsv"

echo "=== Generating stage traces from $(basename "$EXCEL") ==="
"$PYTHON" "$SCRIPT_DIR/ptrace_gen.py" \
  --input-xlsx "$EXCEL" \
  --lcf "$DESIGN/example.lcf" \
  --output-dir "$PTRACE_DIR" \
  --output-prefill "$PREFILL_TRACE" \
  --output-decode "$DECODE_TRACE" \
  --output-time "$TIME_PREFIX" \
  --sampling-intvl "$SAMPLING_INTVL" \
  --decode-count "$DECODE_COUNT"

GRID_ROWS="$(awk '$1 == "-grid_rows" {print $2}' "$DESIGN/example.config")"
GRID_COLS="$(awk '$1 == "-grid_cols" {print $2}' "$DESIGN/example.config")"
GRID_ROWS="${GRID_ROWS:-64}"
GRID_COLS="${GRID_COLS:-64}"

run_stage() {
  local stage="$1"
  local trace="$2"
  local time_file="$3"
  local out_dir="$OUT_BASE/$stage"
  mkdir -p "$out_dir"
  cp "$time_file" "$UNITS_FILE" "$out_dir/"

  echo "=== Running $stage transient: $(basename "$trace") ==="
  # The two HotSpot invocations are independent and each starts at init_temp.
  # HotSpot's transient trace file is append-only, so clear this stage's file.
  : > "$out_dir/example.grid.ttrace"
  (
    cd "$DESIGN"
    "$HOTSPOT" -c example.config \
      -p "$trace" -grid_layer_file example.lcf \
      -materials_file example.materials -model_type grid -detailed_3D on \
      -sampling_intvl "$SAMPLING_INTVL" -init_file '(null)' \
      -steady_file "$out_dir/example.steady" \
      -grid_steady_file "$out_dir/example.grid.steady" \
      -o "$out_dir/example.ttrace" \
      -grid_transient_file "$out_dir/example.grid.ttrace"
  )

  local layer_count
  layer_count="$(awk '$1 == "Layer" {n++} END {print n+0}' "$out_dir/example.grid.steady")"
  if [[ "$layer_count" -eq 0 ]]; then
    echo "error: no Layer headers in $out_dir/example.grid.steady" >&2
    return 1
  fi
  "$PYTHON" "$SCRIPT_DIR/split_grid_steady.py" \
    "$out_dir/example.grid.steady" "$layer_count" "$GRID_ROWS" "$GRID_COLS"

  while IFS=$'\t' read -r layer flp; do
    local grid_file="$out_dir/example_layer${layer}.grid.steady"
    if [[ "$flp" != /* && "$flp" != [A-Za-z]:* ]]; then
      flp="$DESIGN/$flp"
    fi
    "$PYTHON" "$SCRIPT_DIR/grid_thermal_map.py" \
      "$flp" "$grid_file" "$GRID_ROWS" "$GRID_COLS" \
      "$out_dir/layer${layer}_$(basename "$flp" .flp).png"
    if command -v perl >/dev/null 2>&1; then
      perl "$SCRIPT_DIR/grid_thermal_map.pl" \
        "$flp" "$grid_file" "$GRID_ROWS" "$GRID_COLS" \
        > "$out_dir/layer${layer}_$(basename "$flp" .flp).svg"
    fi
  done < <(awk -F '\t' 'NR > 1 && !seen[$2]++ {print $2 "\t" $3}' "$UNITS_FILE")
}

run_stage prefill "$PREFILL_TRACE" "${TIME_PREFIX%.tsv}_prefill.tsv"
run_stage decode "$DECODE_TRACE" "${TIME_PREFIX%.tsv}_decode.tsv"

echo "=== Done. Stage outputs: $OUT_BASE/prefill and $OUT_BASE/decode ==="
