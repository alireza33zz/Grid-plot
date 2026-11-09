"""
    GridPlot

Render single-line diagrams of a distribution feeder from an OpenDSS graph
export. Each diagram (which buses to highlight, how to label them, colours,
sizes) is described in a TOML file under `diagrams/`; the code here is generic.
"""
module GridPlot

using Plots
using TOML
using XLSX

export Network, load_network, bus_position, feeder_position, segments
export DiagramSpec, load_spec, load_specs
export render, save_diagram, validate

include("network.jl")
include("spec.jl")
include("render.jl")
include("cli.jl")

end
