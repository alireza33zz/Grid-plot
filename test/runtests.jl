using GridPlot
using GridPlot: SpecError, UsageError, parse_args, parse_spec, DEFAULT_DATA, DEFAULT_DIAGRAMS
using Test

const NET = load_network(DEFAULT_DATA)
const SPECS = load_specs(DEFAULT_DIAGRAMS)

spec(toml::AbstractString) = parse_spec("test", GridPlot.TOML.parse(toml))

@testset "GridPlot" begin
    @testset "network" begin
        @test length(NET) == 906
        @test bus_position(NET, 34) == (390885.0, 392880.0)
        @test feeder_position(NET, 2) == (390872.66, 392887.38)
        @test_throws KeyError bus_position(NET, 10_000)

        xs, ys = segments(NET)
        @test length(xs) == length(ys) == 3 * (length(NET) - 1)   # bus 1 is a (0,0) placeholder
        @test all(isnan, xs[3:3:end])
        @test_throws ArgumentError load_network("does-not-exist.xlsx")
        @test_throws ArgumentError Network([1, 1], zeros(2), zeros(2), zeros(2), zeros(2))
    end

    @testset "shipped diagrams" begin
        @test !isempty(SPECS)
        @test allunique(s.output for s in SPECS)
        for s in SPECS
            @test validate(s, NET) === nothing
        end
    end

    @testset "spec parsing" begin
        s = spec("""
            title = "T"
            [[markers]]
            buses = [3, 4]
            [[labels]]
            scale = 2
            text = { 3 = "three" }
            offsets = { 3 = [1, -1], 4 = [0.5, 0] }
            """)
        @test s.output == "test" && s.formats == ["png"]
        @test s.markers[1].buses == [3, 4]
        labels = s.labels[1].labels
        @test [(l.bus, l.text, l.offset) for l in labels] == [(3, "three", (2.0, -2.0)), (4, "4", (1.0, 0.0))]

        @test_throws ArgumentError spec("description = \"no title\"")
        @test_throws ArgumentError spec("title = \"T\"\nformats = [\"jpg\"]")
        @test_throws ArgumentError spec("title = \"T\"\ncolour = \"red\"")              # typo'd key
        @test_throws ArgumentError spec("title = \"T\"\n[[labels]]\noffsets = { x = [1, 2] }")
        @test_throws ArgumentError spec("title = \"T\"\n[[labels]]\noffsets = { 3 = [1] }")
        @test_throws ArgumentError spec("title = \"T\"\n[[labels]]\noffsets = {}\ntext = { 3 = \"a\" }")
        @test_throws ArgumentError validate(spec("title = \"T\"\n[[markers]]\nbuses = [99999]"), NET)

        mktemp() do path, io
            write(io, "title = \"T\"\n[lines]\nwidht = 3\n")
            close(io)
            @test_throws SpecError load_spec(path)
        end
    end

    @testset "command line" begin
        o = parse_args(["zones", "-f", "PNG, svg", "--outdir", "out", "--show"])
        @test o.names == ["zones"] && o.formats == ["png", "svg"] && o.outdir == "out" && o.show
        @test parse_args(["--all"]).all
        @test_throws UsageError parse_args(["--bogus"])
        @test_throws UsageError parse_args(["--format"])
        @test_throws UsageError parse_args(["-f", "gif"])
        @test GridPlot.main(["no-such-diagram"]) == 2
    end

    @testset "render" begin
        s = only(filter(s -> s.id == "zone_representatives", SPECS))
        p = render(s, NET)
        mktempdir() do dir
            paths = save_diagram(p, s; outdir = dir, formats = ["png"])
            @test paths == [joinpath(dir, "$(s.output).png")]
            @test filesize(only(paths)) > 0
        end
    end
end
