"""
    validate(spec, net)

Throw an `ArgumentError` listing every bus `spec` refers to that is not in `net`.
"""
function validate(spec::DiagramSpec, net::Network)
    missing_buses = filter(b -> !hasbus(net, b), referenced_buses(spec))
    isempty(missing_buses) ||
        throw(ArgumentError("diagram \"$(spec.id)\" refers to buses not in the network: $(join(sort!(missing_buses), ", "))"))
    return nothing
end

"""
    render(spec, net) -> Plots.Plot

Draw the single-line diagram described by `spec` on top of `net`.
"""
function render(spec::DiagramSpec, net::Network)
    validate(spec, net)
    p = new_canvas(spec.canvas)
    draw_lines!(p, net, spec.lines)
    draw_substation!(p, net, spec.substation)
    foreach(layer -> draw_markers!(p, net, layer), spec.markers)
    foreach(layer -> draw_labels!(p, net, layer), spec.labels)
    return p
end

function new_canvas(c::Canvas)
    margins = Dict(Symbol(side, :_margin) => cm * Plots.Measures.cm for (side, cm) in c.margins_cm)
    font = c.font === nothing ? (;) : (; fontfamily = c.font)
    return plot(; legend = false, grid = false, framestyle = :none,
                size = c.size, dpi = c.dpi, font..., margins...)
end

function draw_lines!(p, net::Network, style::LineStyle)
    xs, ys = segments(net)
    plot!(p, xs, ys; lw = style.width, color = style.color, alpha = style.alpha, label = false)
end

function draw_substation!(p, net::Network, s::Substation)
    x, y = feeder_position(net, s.bus)
    annotate!(p, [
        (x, y + s.marker_dy, text("*", s.marker_color, s.marker_size)),
        (x, y + s.label_dy, text(s.label, s.label_color, s.label_size)),
    ])
end

function draw_markers!(p, net::Network, layer::MarkerLayer)
    isempty(layer.buses) && return p
    glyph = text("*", layer.color, layer.size)
    annotate!(p, [(x, y + layer.dy, glyph) for (x, y) in bus_position.(Ref(net), layer.buses)])
end

function draw_labels!(p, net::Network, layer::LabelLayer)
    isempty(layer.labels) && return p
    annotate!(p, map(layer.labels) do l
        x, y = bus_position(net, l.bus)
        (x + l.offset[1], y + l.offset[2], text(l.text, layer.color, layer.size))
    end)
end

"""
    save_diagram(p, spec; outdir, formats = spec.formats) -> Vector{String}

Write `p` as `<outdir>/<spec.output>.<fmt>` for each format and return the paths.
"""
function save_diagram(p, spec::DiagramSpec; outdir::AbstractString,
                      formats::AbstractVector{<:AbstractString} = spec.formats)
    mkpath(outdir)
    paths = [joinpath(outdir, "$(spec.output).$fmt") for fmt in formats]
    foreach(path -> savefig(p, path), paths)
    return paths
end
