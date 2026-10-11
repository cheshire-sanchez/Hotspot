#!/usr/bin/env bash
set -euo pipefail

# Usage: bash scripts/run.sh [design-directory] [Excel-file-in-psource]
# The default workbook is psource/W4_power_2.xlsx.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DESIGN="${1:-$REPO_ROOT/HBFuse-2}"
DESIGN="$(cd "$DESIGN" && pwd)"
DESIGN_NAME="$(basename "$DESIGN")"
SOURCE_DIR="$REPO_ROOT/psource"
PTRACE_DIR="$REPO_ROOT/ptrace"
OUT_DIR="$REPO_ROOT/outputs/$DESIGN_NAME"
HOTSPOT="${HOTSPOT:-$REPO_ROOT/hotspot}"
SAMPLING_INTVL=0.1
DECODE_COUNT=129

if command -v python3 >/dev/null 2>&1; then
  PYTHON=python3
elif command -v python >/dev/null 2>&1; then
  PYTHON=python
else
  echo "error: Python not found" >&2
  exit 1
fi

EXCEL="$SOURCE_DIR/$(basename "${2:-W4_power_2.xlsx}")"
for file in "$EXCEL" "$DESIGN/example.lcf" "$DESIGN/example.config" "$DESIGN/example.materials" "$HOTSPOT"; do
  if [[ ! -f "$file" ]]; then
    echo "error: file not found: $file" >&2
    exit 1
  fi
done

mkdir -p "$PTRACE_DIR" "$OUT_DIR"
PTRACE_FILE="$PTRACE_DIR/$DESIGN_NAME.ptrace"
TIME_FILE="$PTRACE_DIR/$DESIGN_NAME.time.tsv"
UNITS_FILE="$PTRACE_DIR/$DESIGN_NAME.units.tsv"

echo "=== Generating $DESIGN_NAME trace from $(basename "$EXCEL") ==="
"$PYTHON" "$SCRIPT_DIR/ptrace_gen.py" \
  --input-xlsx "$EXCEL" \
  --lcf "$DESIGN/example.lcf" \
  --output-dir "$PTRACE_DIR" \
  --sampling-intvl "$SAMPLING_INTVL" \
  --decode-count "$DECODE_COUNT"
cp "$TIME_FILE" "$UNITS_FILE" "$OUT_DIR/"

# Override config at runtime so trace generation and HotSpot use the same dt.
GRID_ROWS="$(awk '$1 == "-grid_rows" {print $2}' "$DESIGN/example.config")"
GRID_COLS="$(awk '$1 == "-grid_cols" {print $2}' "$DESIGN/example.config")"
GRID_ROWS="${GRID_ROWS:-64}"
GRID_COLS="${GRID_COLS:-64}"

echo "=== Computing steady temperatures: $DESIGN_NAME ==="
(
  cd "$DESIGN"
  "$HOTSPOT" -c example.config \
    -p "$PTRACE_FILE" -grid_layer_file example.lcf \
    -materials_file example.materials -model_type grid -detailed_3D on \
    -sampling_intvl "$SAMPLING_INTVL" \
    -steady_file "$OUT_DIR/example.steady" \
    -grid_steady_file "$OUT_DIR/example.grid.steady"
)

# Preserve the actual prefill-to-decode warm-up, starting at config init_temp.
echo "=== Computing transient temperatures: $DESIGN_NAME ==="
# HotSpot appends grid snapshots; start a new output on each run.
: > "$OUT_DIR/example.grid.ttrace"
(
  cd "$DESIGN"
  "$HOTSPOT" -c example.config \
    -p "$PTRACE_FILE" -grid_layer_file example.lcf \
    -materials_file example.materials -model_type grid -detailed_3D on \
    -sampling_intvl "$SAMPLING_INTVL" -init_file '(null)' \
    -o "$OUT_DIR/example.ttrace" \
    -grid_transient_file "$OUT_DIR/example.grid.ttrace"
)

# Count actual grid output layers, including HotSpot's package layers.
LAYER_COUNT="$(awk '$1 == "Layer" {n++} END {print n+0}' "$OUT_DIR/example.grid.steady")"
if [[ "$LAYER_COUNT" -eq 0 ]]; then
  echo "error: no Layer headers in grid steady output" >&2
  exit 1
fi
"$PYTHON" "$SCRIPT_DIR/split_grid_steady.py" \
  "$OUT_DIR/example.grid.steady" "$LAYER_COUNT" "$GRID_ROWS" "$GRID_COLS"

# The generated unit map provides the ordered power layers and their FLPs.
while IFS=$'\t' read -r layer flp; do
  GRID_FILE="$OUT_DIR/example_layer${layer}.grid.steady"
  if [[ "$flp" != /* && "$flp" != [A-Za-z]:* ]]; then
    flp="$DESIGN/$flp"
  fi
  "$PYTHON" "$SCRIPT_DIR/grid_thermal_map.py" \
    "$flp" "$GRID_FILE" "$GRID_ROWS" "$GRID_COLS" \
    "$OUT_DIR/layer${layer}_$(basename "$flp" .flp).png"
  if command -v perl >/dev/null 2>&1; then
    perl "$SCRIPT_DIR/grid_thermal_map.pl" \
      "$flp" "$GRID_FILE" "$GRID_ROWS" "$GRID_COLS" \
      > "$OUT_DIR/layer${layer}_$(basename "$flp" .flp).svg"
  fi
done < <(awk -F '\t' 'NR > 1 && !seen[$2]++ {print $2 "\t" $3}' "$UNITS_FILE")

echo "=== Done. Trace: $PTRACE_FILE; results: $OUT_DIR ==="
