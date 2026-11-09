"""
    Network

Feeder geometry as drawn on the single-line diagram.

Every bus `b` is fed by one line segment running from `(x1, y1)` (the upstream
bus) to `(x2, y2)` (bus `b` itself). Buses are addressed by the numbers in the
spreadsheet's `Bus Number` column, so diagram definitions never depend on row
order.
"""
struct Network
    buses::Vector{Int}
    x1::Vector{Float64}
    y1::Vector{Float64}
    x2::Vector{Float64}
    y2::Vector{Float64}
    row::Dict{Int,Int}   # bus number => index into the coordinate vectors
end

function Network(buses, x1, y1, x2, y2)
    row = Dict{Int,Int}()
    for (i, b) in pairs(buses)
        haskey(row, b) && throw(ArgumentError("bus $b appears more than once"))
        row[b] = i
    end
    return Network(buses, x1, y1, x2, y2, row)
end

Base.length(net::Network) = length(net.buses)
Base.show(io::IO, net::Network) = print(io, "Network(", length(net), " buses)")

hasbus(net::Network, bus::Integer) = haskey(net.row, bus)

function _row(net::Network, bus::Integer)
    i = get(net.row, bus, nothing)
    i === nothing && throw(KeyError(bus))
    return i
end

"Position of `bus` (the downstream end of its feeding segment)."
bus_position(net::Network, bus::Integer) = (i = _row(net, bus); (net.x2[i], net.y2[i]))

"Position of the upstream end of the segment feeding `bus`."
feeder_position(net::Network, bus::Integer) = (i = _row(net, bus); (net.x1[i], net.y1[i]))

"""
    segments(net) -> (xs, ys)

All drawable line segments packed into two vectors separated by `NaN`, so the
whole feeder can be drawn as a single series. Placeholder rows whose four
coordinates are all zero (the source bus in OpenDSS exports) are skipped.
"""
function segments(net::Network)
    xs = Float64[]
    ys = Float64[]
    for i in eachindex(net.buses)
        coords = (net.x1[i], net.y1[i], net.x2[i], net.y2[i])
        all(iszero, coords) && continue
        append!(xs, (net.x1[i], net.x2[i], NaN))
        append!(ys, (net.y1[i], net.y2[i], NaN))
    end
    return xs, ys
end

"""
    load_network(path; sheet = nothing) -> Network

Read an OpenDSS graph export with columns `Bus Number, X1, Y1, X2, Y2` and a
single header row. Uses the first sheet unless `sheet` is given.
"""
function load_network(path::AbstractString; sheet::Union{Nothing,AbstractString} = nothing)
    isfile(path) || throw(ArgumentError("data file not found: $path"))
    table = XLSX.openxlsx(path) do xf
        name = something(sheet, first(XLSX.sheetnames(xf)))
        xf[name][:]
    end
    size(table, 2) >= 5 ||
        throw(ArgumentError("$path: expected 5 columns (Bus Number, X1, Y1, X2, Y2), found $(size(table, 2))"))

    body = table[2:end, 1:5]                          # drop the header row
    body = body[.!all(ismissing, body; dims = 2)[:], :]  # drop trailing blank rows
    any(ismissing, body) && throw(ArgumentError("$path: data contains empty cells"))

    col(j) = Float64.(body[:, j])
    return Network(Int.(body[:, 1]), col(2), col(3), col(4), col(5))
end
