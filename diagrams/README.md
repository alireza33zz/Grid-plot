# Diagram definitions

Each `*.toml` file in this folder is one figure; the file name (without
`.toml`) is the id used on the command line. Every section except `title` is
optional and falls back to the defaults shown below. Unknown keys are rejected,
so typos are caught instead of silently ignored.

Coordinates and offsets are in the same units as the spreadsheet (metres in the
shipped data). Text and marker sizes are font points.

```toml
title       = "Three zones"                    # required; shown in the menu
description = "One-line summary for --list"
output      = "SingleLineDiagramZones"         # file name without extension (default: the id)
formats     = ["png"]                          # any of "png", "pdf", "svg"

[canvas]
size       = [6300, 4000]                      # pixels at 100 dpi
dpi        = 100
font       = "Courier"                         # optional font family
margins_cm = { top = 2 }                       # any of top, bottom, left, right

[lines]                                        # the feeder itself
width = 15
color = "blue"                                 # any Plots.jl colour name
alpha = 0.9

[substation]
bus    = 2                                     # drawn at the upstream end of this bus's line
marker = { color = "red", size = 150, dy = -1.25 }
label  = { text = "Substation", color = "red", size = 90, dy = 5 }

# Highlight a group of buses with a "*". Repeat the block for more groups;
# they are drawn in file order.
[[markers]]
name  = "Zone 1"                               # for readability only
buses = [25, 32, 59]
color = "magenta"
size  = 120
dy    = -1.0                                   # nudges the glyph so it centres on the bus

# Text labels. The buses listed under `offsets` are the ones that get a label.
[[labels]]
name  = "Bus numbers"
color = "black"
size  = 80
scale = 1.0                                    # multiplies every offset

[labels.offsets]                               # bus = [dx, dy] from the bus position
25 = [-1, -4]
32 = [-1, -4]

[labels.text]                                  # optional; default text is the bus number
25 = "Z1.1"
```
