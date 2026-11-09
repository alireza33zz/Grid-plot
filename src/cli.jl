const PROJECT_ROOT = normpath(joinpath(@__DIR__, ".."))
const DEFAULT_DATA = joinpath(PROJECT_ROOT, "data", "DSSGraph_Output.xlsx")
const DEFAULT_DIAGRAMS = joinpath(PROJECT_ROOT, "diagrams")
const DEFAULT_OUTDIR = joinpath(PROJECT_ROOT, "output")

const USAGE = """
Usage: julia gridplot.jl [options] [DIAGRAM ...]

Render single-line diagrams of the feeder. With no DIAGRAM and no --all, you
are asked step by step which diagrams to render and in which format.

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
    parse_choice(reply, n; multiple) -> Union{Int, Vector{Int}, Nothing}

Interpret a typed answer to a numbered question with `n` options: a single
number, or with `multiple = true` a comma/space-separated list of numbers or
`a`/`all`. Returns `nothing` if the answer is not valid.
"""
function parse_choice(reply::AbstractString, n::Integer; multiple::Bool)
    multiple && lowercase(reply) in ("a", "all") && return collect(1:n)
    picks = [tryparse(Int, t) for t in split(reply, r"[,\s]+"; keepempty = false)]
    (isempty(picks) || any(p -> p === nothing || !(1 <= p <= n), picks)) && return nothing
    multiple && return unique!(Vector{Int}(picks))
    return length(picks) == 1 ? only(picks) : nothing
end

"""
    ask(question, options; multiple = false, default = nothing, input = stdin, output = stdout)

Print `question` with numbered `options` and read answers until a valid one is
given (see [`parse_choice`](@ref)). An empty answer selects `default` if set.
Returns `nothing` if the user types `q` or input ends.
"""
function ask(question::AbstractString, options::AbstractVector{<:AbstractString};
             multiple::Bool = false, default::Union{Nothing,Integer} = nothing,
             input::IO = stdin, output::IO = stdout)
    println(output)
    printstyled(output, question, "\n"; bold = true)
    width = ndigits(length(options))
    for (i, opt) in enumerate(options)
        println(output, "  ", lpad(i, width), ") ", opt)
    end
    hint = multiple ? "Numbers separated by commas, or 'a' for all" : "Number"
    hint *= default === nothing ? "" : " [default $default]"
    while true
        print(output, hint, " ('q' to quit): ")
        eof(input) && (println(output); return nothing)
        reply = strip(readline(input))
        isempty(reply) && default !== nothing && (reply = string(default))
        lowercase(reply) in ("q", "quit") && return nothing
        choice = parse_choice(reply, length(options); multiple)
        choice === nothing || return choice
        printstyled(output, "  \"", reply, "\" is not a valid choice, please try again.\n"; color = :yellow)
    end
end

"""
Ask, one question at a time, which diagrams to render, in which formats, and
whether to preview them. Questions already answered on the command line
(`formats`, `show`) are skipped. Returns `(; specs, formats, show)`, with
`formats === nothing` meaning "use each diagram's own formats", or `nothing`
if the user quit.
"""
function interactive_selection(specs; formats = nothing, show::Bool = false,
                               input::IO = stdin, output::IO = stdout)
    printstyled(output, "Grid Plot — single-line diagram renderer\n"; bold = true, color = :blue)

    picks = ask("1. Which diagrams do you want to render?",
                ["$(s.title)  ($(s.id))" for s in specs]; multiple = true, input, output)
    picks === nothing && return nothing
    selected = specs[sort(picks)]

    if formats === nothing
        options = ["Each diagram's default", "PNG", "PDF", "SVG", "All formats (PNG, PDF, SVG)"]
        pick = ask("2. Which output format?", options; default = 1, input, output)
        pick === nothing && return nothing
        formats = pick == 1 ? nothing :
                  pick == 5 ? collect(SUPPORTED_FORMATS) : [lowercase(options[pick])]
    end

    if !show
        pick = ask("3. Open a preview window after saving each diagram?", ["No", "Yes"];
                   default = 1, input, output)
        pick === nothing && return nothing
        show = pick == 2
    end

    println(output)
    return (; specs = selected, formats, show)
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

        formats, show = opts.formats, opts.show
        if opts.all
            selected = specs
        elseif !isempty(opts.names)
            selected = select_specs(specs, opts.names)
        else
            choice = interactive_selection(specs; formats, show)
            choice === nothing && (println("Nothing selected."); return 0)
            (; formats, show) = choice
            selected = choice.specs
        end

        net = load_network(opts.data)
        foreach(s -> validate(s, net), selected)   # fail fast before any slow rendering
        run_diagrams(selected, net; outdir = opts.outdir, formats, show)
        return 0
    catch err
        err isa Union{UsageError,SpecError,ArgumentError} || rethrow()
        printstyled(stderr, "error: "; color = :red, bold = true)
        println(stderr, sprint(showerror, err))
        return err isa UsageError ? 2 : 1
    end
end
