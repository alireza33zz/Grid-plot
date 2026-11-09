# Grid Plot

Single-line diagrams of a distribution feeder, rendered from an OpenDSS graph
export with [Plots.jl](https://docs.juliaplots.org).

Each diagram is described in a small TOML file under [`diagrams/`](diagrams/):
which buses to highlight, how to label them, and the colours and sizes to use.
To add or tweak a figure, edit a TOML file. You don't need to touch the code.

## Quick start

```sh
julia --project=. -e "using Pkg; Pkg.instantiate()"   # first time only
julia gridplot.jl                                      # asks what to render
```

You are asked, one question at a time, which diagrams to render (type numbers
like `1,3`, or `a` for all), which output format to use, and whether to open a
preview. Press Enter to accept a default, or `q` to quit.
Figures are written to `output/`.

### Command line

```sh
julia gridplot.jl --list                     # show available diagrams
julia gridplot.jl zones                      # render one diagram
julia gridplot.jl --all -f png,pdf           # everything, as PNG and PDF
julia gridplot.jl numbered_sites --show      # render, then open a preview window
julia gridplot.jl --help
```

| Option              | Meaning                                                   |
| ------------------- | --------------------------------------------------------- |
| `-a`, `--all`       | Render every diagram                                      |
| `-l`, `--list`      | List diagrams and exit                                    |
| `-f`, `--format`    | `png`, `pdf`, `svg` (comma-separated); default per diagram |
| `-o`, `--outdir`    | Output directory (default `output/`)                      |
| `-d`, `--data`      | Network spreadsheet (default `data/DSSGraph_Output.xlsx`) |
| `--diagrams`        | Directory of diagram definitions (default `diagrams/`)    |
| `--show`            | Open each figure in a window after saving                 |

### From Julia

```julia
using GridPlot
net  = load_network("data/DSSGraph_Output.xlsx")
spec = load_spec("diagrams/zones.toml")
p    = render(spec, net)                 # a regular Plots.jl plot
save_diagram(p, spec; outdir = "output", formats = ["svg"])
```

## Diagrams

| Id                     | Output file              | Shows                                                    |
| ---------------------- | ------------------------ | -------------------------------------------------------- |
| `bus_labels`           | `SingleLineDiagram`      | The 55 candidate buses, labelled with bus numbers        |
| `numbered_sites`       | `SingleLineDiagram4`     | Sites numbered 1–55, rooftop PV in red, DER1–DER3        |
| `zones`                | `SingleLineDiagramZones` | Three colour-coded zones of eleven buses each            |
| `zone_representatives` | `SingleLineDiagramLines` | Three representative buses per zone (Z1.1 … Z3.3)        |

See [`diagrams/README.md`](diagrams/README.md) for the file format.

## Layout

```
gridplot.jl          entry point (menu + CLI)
src/
  GridPlot.jl        module definition
  network.jl         reading the spreadsheet, bus coordinates
  spec.jl            diagram definition types and TOML parsing/validation
  render.jl          drawing and saving
  cli.jl             argument parsing and the interactive menu
diagrams/            one TOML file per figure
data/                OpenDSS export (Bus Number, X1, Y1, X2, Y2)
figures/legacy/      figures produced by the original scripts, kept for reference
test/                run with: julia --project=. -e "using Pkg; Pkg.test()"
```

## Input data

`data/DSSGraph_Output.xlsx` has one header row followed by one row per bus:
`Bus Number, X1, Y1, X2, Y2`. Bus *b* sits at `(X2, Y2)` and is fed by the line
from `(X1, Y1)`. Rows with all-zero coordinates (the source placeholder) are
not drawn.
