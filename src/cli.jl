const PROJECT_ROOT = normpath(joinpath(@__DIR__, ".."))
const DEFAULT_DATA = joinpath(PROJECT_ROOT, "data", "DSSGraph_Output.xlsx")
const DEFAULT_DIAGRAMS = joinpath(PROJECT_ROOT, "diagrams")
const DEFAULT_OUTDIR = joinpath(PROJECT_ROOT, "output")

const USAGE = """
Usage: julia gridplot.jl [options] [DIAGRAM ...]

Render single-line diagrams of the feeder. With no DIAGRAM and no --all, an
interactive menu lets you pick diagrams and output formats.

Options:
  -a, --all             Render every diagram
  -l, --list            List available diagrams and exit
  -f, --format FMTS     Comma-separated output formats (png,pdf,svg);
                        default: the formats set in each diagram file
  -o, --outdir DIR      Output directory (default: output/)
  -d, --data FILE       Network spreadsheet (default: data/DSSGraph_Output.xlsx)
      --diagrams DIR    Directory of diagram definitions (default: diagrams/)
      --show            Open each diagram in a window after saving
  -h, --help            Show this help
"""

Base.@kwdef mutable struct Options
    names::Vector{String} = String[]
    all::Bool = false
    list::Bool = false
    help::Bool = false
    show::Bool = false
    formats::Union{Nothing,Vector{String}} = nothing
    outdir::String = DEFAULT_OUTDIR
    data::String = DEFAULT_DATA
    diagrams::String = DEFAULT_DIAGRAMS
end

struct UsageError <: Exception
    msg::String
end
Base.showerror(io::IO, e::UsageError) = print(io, e.msg)

function parse_args(args::AbstractVector{<:AbstractString})
    opts = Options()
    i = firstindex(args)
    function value(flag)
        i += 1
        i <= lastindex(args) || throw(UsageError("option $flag requires a value"))
        return args[i]
    end
    while i <= lastindex(args)
        arg = args[i]
        if arg in ("-h", "--help")
            opts.help = true
        elseif arg in ("-a", "--all")
            opts.all = true
        elseif arg in ("-l", "--list")
            opts.list = true
        elseif arg == "--show"
            opts.show = true
        elseif arg in ("-f", "--format")
            opts.formats = parse_formats(value(arg))
        elseif arg in ("-o", "--outdir")
            opts.outdir = value(arg)
        elseif arg in ("-d", "--data")
            opts.data = value(arg)
        elseif arg == "--diagrams"
            opts.diagrams = value(arg)
        elseif startswith(arg, "-")
            throw(UsageError("unknown option $arg"))
        else
            push!(opts.names, arg)
        end
        i += 1
    end
    return opts
end

function parse_formats(s::AbstractString)
    formats = unique!(lowercase.(strip.(split(s, ','; keepempty = false))))
    isempty(formats) && throw(UsageError("--format needs at least one format"))
    for f in formats
        f in SUPPORTED_FORMATS ||
            throw(UsageError("unsupported format \"$f\" (supported: $(join(SUPPORTED_FORMATS, ", ")))"))
    end
    return formats
end

function print_list(io::IO, specs)
    width = maximum(s -> length(s.id), specs; init = 0)
    for s in specs
        println(io, "  ", rpad(s.id, width), "  ", s.title)
        isempty(s.description) || println(io, "  ", " "^width, "  ", s.description)
    end
end

function select_specs(specs, names)
    by_id = Dict(s.id => s for s in specs)
    unknown = filter(n -> !haskey(by_id, n), names)
    isempty(unknown) ||
        throw(UsageError("unknown diagram(s): $(join(unknown, ", ")). Run with --list to see the options."))
    return [by_id[n] for n in unique(names)]
end

"""
Interactive selection. Returns `(specs, formats)`, where `formats === nothing`
means "use each diagram's own formats", or `nothing` if the user cancelled.
"""
function interactive_selection(specs)
    menu = TerminalMenus.MultiSelectMenu(["$(s.id) — $(s.title)" for s in specs]; pagesize = 10)
    println("Grid Plot — single-line diagram renderer")
    chosen = TerminalMenus.request(
        "Select diagrams ([space] toggle, [a] all, [enter] confirm, [q] quit):", menu)
    isempty(chosen) && return nothing

    choices = ["Diagram defaults", "PNG", "PDF", "SVG", "PNG + PDF + SVG"]
    pick = TerminalMenus.request("Output format:", TerminalMenus.RadioMenu(choices))
    pick == -1 && return nothing
    formats = pick == 1 ? nothing :
              pick == 5 ? collect(SUPPORTED_FORMATS) : [lowercase(choices[pick])]
    return specs[sort!(collect(chosen))], formats
end

function run_diagrams(specs, net::Network; outdir, formats, show::Bool, io::IO = stdout)
    for (k, spec) in enumerate(specs)
        print(io, "[", k, "/", length(specs), "] ", spec.id, " … ")
        flush(io)
        t = @elapsed begin
            p = render(spec, net)
            paths = save_diagram(p, spec; outdir, formats = something(formats, spec.formats))
        end
        println(io, "done (", round(t; digits = 1), " s)")
        foreach(path -> println(io, "      → ", relpath(path)), paths)
        if show
            display(p)
            print(io, "      Press Enter to continue…")
            readline()
        end
    end
end

"""
    main(args = ARGS) -> Int

Command-line entry point; returns a process exit code.
"""
function main(args::AbstractVector{<:AbstractString} = ARGS)
    try
        opts = parse_args(args)
        if opts.help
            print(USAGE)
            return 0
        end

        specs = load_specs(opts.diagrams)
        isempty(specs) && throw(UsageError("no diagram definitions found in $(opts.diagrams)"))
        if opts.list
            print_list(stdout, specs)
            return 0
        end

        formats = opts.formats
        if opts.all
            selected = specs
        elseif !isempty(opts.names)
            selected = select_specs(specs, opts.names)
        elseif stdin isa Base.TTY
            choice = interactive_selection(specs)
            choice === nothing && (println("Nothing selected."); return 0)
            selected, menu_formats = choice
            formats = something(opts.formats, menu_formats, Some(nothing))
        else
            throw(UsageError("no diagram selected; pass diagram names or --all (see --help)"))
        end

        net = load_network(opts.data)
        foreach(s -> validate(s, net), selected)   # fail fast before any slow rendering
        run_diagrams(selected, net; outdir = opts.outdir, formats, show = opts.show)
        return 0
    catch err
        err isa Union{UsageError,SpecError,ArgumentError} || rethrow()
        printstyled(stderr, "error: "; color = :red, bold = true)
        println(stderr, sprint(showerror, err))
        return err isa UsageError ? 2 : 1
    end
end
