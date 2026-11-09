const SUPPORTED_FORMATS = ("png", "pdf", "svg")

const Offset = NTuple{2,Float64}

"Canvas-level settings passed straight to `Plots.plot`."
Base.@kwdef struct Canvas
    size::NTuple{2,Int} = (6300, 4000)
    dpi::Int = 100
    font::Union{Nothing,String} = nothing
    margins_cm::Dict{Symbol,Float64} = Dict(:top => 2.0)
end

Base.@kwdef struct LineStyle
    width::Float64 = 15
    color::Symbol = :blue
    alpha::Float64 = 0.9
end

"""
A group of buses highlighted with the same `*` glyph. `dy` nudges the glyph
down so its visual centre sits on the bus (text is anchored at the baseline).
"""
Base.@kwdef struct MarkerLayer
    name::String = ""
    buses::Vector{Int}
    color::Symbol = :red
    size::Int = 60
    dy::Float64 = -0.5
end

"One text label: `text` drawn at the bus position shifted by `offset`."
struct Label
    bus::Int
    text::String
    offset::Offset
end

Base.@kwdef struct LabelLayer
    name::String = ""
    labels::Vector{Label}
    color::Symbol = :black
    size::Int = 80
end

Base.@kwdef struct Substation
    bus::Int = 2
    marker_color::Symbol = :red
    marker_size::Int = 150
    marker_dy::Float64 = -1.25
    label::String = "Substation"
    label_color::Symbol = :red
    label_size::Int = 90
    label_dy::Float64 = 5
end

"Everything needed to draw one single-line diagram."
Base.@kwdef struct DiagramSpec
    id::String
    title::String
    description::String = ""
    output::String
    formats::Vector{String} = ["png"]
    canvas::Canvas = Canvas()
    lines::LineStyle = LineStyle()
    substation::Substation = Substation()
    markers::Vector{MarkerLayer} = MarkerLayer[]
    labels::Vector{LabelLayer} = LabelLayer[]
end

Base.show(io::IO, s::DiagramSpec) = print(io, "DiagramSpec(\"", s.id, "\")")

"All buses a spec refers to, for validation against a `Network`."
referenced_buses(s::DiagramSpec) = unique!(vcat(
    s.substation.bus,
    (m.buses for m in s.markers)...,
    ([l.bus for l in layer.labels] for layer in s.labels)...,
))

# ---------------------------------------------------------------------------
# TOML loading
# ---------------------------------------------------------------------------

struct SpecError <: Exception
    file::String
    msg::String
end
Base.showerror(io::IO, e::SpecError) = print(io, "invalid diagram file ", e.file, ": ", e.msg)

"""
    load_spec(path) -> DiagramSpec

Parse a diagram definition. The file's base name becomes the diagram id.
See `diagrams/README.md` for the format.
"""
function load_spec(path::AbstractString)
    id = splitext(basename(path))[1]
    raw = try
        TOML.parsefile(path)
    catch err
        err isa TOML.ParserError || rethrow()
        throw(SpecError(path, sprint(showerror, err)))
    end
    try
        return parse_spec(id, raw)
    catch err
        err isa SpecError && rethrow()
        err isa Union{ArgumentError,KeyError,MethodError,InexactError} || rethrow()
        throw(SpecError(path, sprint(showerror, err)))
    end
end

"""
    load_specs(dir) -> Vector{DiagramSpec}

Load every `*.toml` file in `dir`, sorted by id.
"""
function load_specs(dir::AbstractString)
    isdir(dir) || throw(ArgumentError("diagram directory not found: $dir"))
    files = filter(f -> endswith(f, ".toml"), readdir(dir; join = true, sort = true))
    return [load_spec(f) for f in files]
end

function parse_spec(id::AbstractString, d::AbstractDict)
    check_keys(d, ("title", "description", "output", "formats", "canvas", "lines",
                   "substation", "markers", "labels"), "top level")
    formats = String.(get(d, "formats", ["png"]))
    for f in formats
        f in SUPPORTED_FORMATS ||
            throw(ArgumentError("unsupported format \"$f\" (supported: $(join(SUPPORTED_FORMATS, ", ")))"))
    end
    return DiagramSpec(;
        id = String(id),
        title = String(required(d, "title", "top level")),
        description = String(get(d, "description", "")),
        output = String(get(d, "output", id)),
        formats,
        canvas = parse_canvas(get(d, "canvas", Dict())),
        lines = parse_lines(get(d, "lines", Dict())),
        substation = parse_substation(get(d, "substation", Dict())),
        markers = [parse_marker(m) for m in get(d, "markers", [])],
        labels = [parse_label_layer(l) for l in get(d, "labels", [])],
    )
end

function parse_canvas(d)
    check_keys(d, ("size", "dpi", "font", "margins_cm"), "[canvas]")
    defaults = Canvas()
    margins = Dict(Symbol(k) => Float64(v) for (k, v) in get(d, "margins_cm", Dict("top" => 2.0)))
    check_keys(margins, (:top, :bottom, :left, :right), "[canvas.margins_cm]")
    return Canvas(;
        size = haskey(d, "size") ? Tuple(Int.(d["size"]))::NTuple{2,Int} : defaults.size,
        dpi = Int(get(d, "dpi", defaults.dpi)),
        font = haskey(d, "font") ? String(d["font"]) : nothing,
        margins_cm = margins,
    )
end

function parse_lines(d)
    check_keys(d, ("width", "color", "alpha"), "[lines]")
    defaults = LineStyle()
    return LineStyle(;
        width = Float64(get(d, "width", defaults.width)),
        color = Symbol(get(d, "color", defaults.color)),
        alpha = Float64(get(d, "alpha", defaults.alpha)),
    )
end

function parse_substation(d)
    check_keys(d, ("bus", "marker", "label"), "[substation]")
    m = get(d, "marker", Dict())
    l = get(d, "label", Dict())
    check_keys(m, ("color", "size", "dy"), "[substation.marker]")
    check_keys(l, ("text", "color", "size", "dy"), "[substation.label]")
    defaults = Substation()
    return Substation(;
        bus = Int(get(d, "bus", defaults.bus)),
        marker_color = Symbol(get(m, "color", defaults.marker_color)),
        marker_size = Int(get(m, "size", defaults.marker_size)),
        marker_dy = Float64(get(m, "dy", defaults.marker_dy)),
        label = String(get(l, "text", defaults.label)),
        label_color = Symbol(get(l, "color", defaults.label_color)),
        label_size = Int(get(l, "size", defaults.label_size)),
        label_dy = Float64(get(l, "dy", defaults.label_dy)),
    )
end

function parse_marker(d)
    check_keys(d, ("name", "buses", "color", "size", "dy"), "[[markers]]")
    defaults = MarkerLayer(; buses = Int[])
    return MarkerLayer(;
        name = String(get(d, "name", "")),
        buses = Int.(required(d, "buses", "[[markers]]")),
        color = Symbol(get(d, "color", defaults.color)),
        size = Int(get(d, "size", defaults.size)),
        dy = Float64(get(d, "dy", defaults.dy)),
    )
end

# Labelled buses are the keys of `offsets`; `text` optionally overrides the
# default label (the bus number) for some or all of them.
function parse_label_layer(d)
    ctx = "[[labels]]"
    check_keys(d, ("name", "color", "size", "scale", "offsets", "text"), ctx)
    defaults = LabelLayer(; labels = Label[])
    scale = Float64(get(d, "scale", 1.0))
    offsets = Dict(parse_bus(k, ctx) => parse_offset(v, k) .* scale
                   for (k, v) in required(d, "offsets", ctx))
    texts = Dict(parse_bus(k, ctx) => string(v) for (k, v) in get(d, "text", Dict()))
    orphans = setdiff(keys(texts), keys(offsets))
    isempty(orphans) ||
        throw(ArgumentError("$ctx: text given for buses without an offset: $(join(sort!(collect(orphans)), ", "))"))
    labels = [Label(bus, get(texts, bus, string(bus)), offsets[bus]) for bus in sort!(collect(keys(offsets)))]
    return LabelLayer(;
        name = String(get(d, "name", "")),
        labels,
        color = Symbol(get(d, "color", defaults.color)),
        size = Int(get(d, "size", defaults.size)),
    )
end

function parse_bus(key::AbstractString, ctx)
    bus = tryparse(Int, key)
    bus === nothing && throw(ArgumentError("$ctx: \"$key\" is not a bus number"))
    return bus
end

function parse_offset(v, key)
    (v isa AbstractVector && length(v) == 2 && all(x -> x isa Real, v)) ||
        throw(ArgumentError("offset for bus $key must be [dx, dy], got $(repr(v))"))
    return (Float64(v[1]), Float64(v[2]))
end

function required(d, key, ctx)
    haskey(d, key) || throw(ArgumentError("$ctx: missing required key \"$key\""))
    return d[key]
end

function check_keys(d, allowed, ctx)
    unknown = setdiff(keys(d), allowed)
    isempty(unknown) ||
        throw(ArgumentError("$ctx: unknown key(s) $(join(repr.(collect(unknown)), ", ")); allowed: $(join(allowed, ", "))"))
    return nothing
end
