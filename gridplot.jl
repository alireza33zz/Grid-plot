#!/usr/bin/env julia
# Entry point: `julia gridplot.jl --help`
import Pkg
Pkg.activate(@__DIR__; io = devnull)

using GridPlot
exit(GridPlot.main(ARGS))
