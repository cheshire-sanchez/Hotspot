# HBFuse-2 HotSpot run

This design uses `example.lcf` to define the stacked layers and power-column order.
The trace generator reads an Excel phase table and writes the generated trace under
the repository `ptrace/` directory. The default source is `psource/W4_power_2.xlsx`.

From the repository root, run in Linux or WSL:

```bash
python3 -m pip install openpyxl
bash scripts/run.sh HBFuse-2 W4_power_2.xlsx
```

The script uses the sampling interval in `HBFuse-2/example.config` (`0.1` seconds),
and passes the same value to HotSpot. Generated simulation outputs are kept under
`outputs/HBFuse-2/` and are ignored by Git.
