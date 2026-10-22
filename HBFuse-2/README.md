# HotSpot workloads

`HBFuse-4/example.lcf` defines the current stacked design and power-column order.
The trace generator reads `psource/W4_power_4.xlsx` and writes separate prefill and
decode traces under `ptrace/`.

On Windows, install MinGW-w64 GCC and Python with `openpyxl`, `numpy`, and
`matplotlib`. From PowerShell at the repository root, run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\run.ps1
```

The script builds `build/windows/hotspot.exe` when needed. It uses the 0.001-second
sampling interval from `HBFuse-4/example.config` and writes prefill and decode
results to separate directories under `outputs/HBFuse-4/`. Build products and
generated simulation data are ignored by Git.
