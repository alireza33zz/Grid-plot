using XLSX
using Plots

ENV["GKSwstype"] = "svg"
ENV["GKS_FILEPATH"] = "SingleLineDiagramLines.svg"
ENV["GKSwstype_DPI"] = "600"

selected_indices = [34, 47, 70, 73, 74, 83, 178, 208, 225, 248, 249, 264, 276, 289, 314, 320,
    327, 337, 342, 349, 387, 388, 406, 458, 502, 522, 539, 556, 562, 563, 611, 614, 619, 
    629, 639, 676, 682, 688, 701, 702, 755, 778, 780, 785, 813, 817, 835, 860, 861, 886,
    896, 898, 899, 900, 906]

    offsets = Dict(
        34 => (3, 0),
        47 => (0, 3),
        70 => (0, 3),
        73 => (-3, 1),
        74 => (3, -2),
        83 => (0, -3),
        178 => (0, 3),
        208 => (-4, 1),
        225 => (-3, 2),
        248 => (3, -3),
        249 => (-3, 2),
        264 => (-3, 3),
        276 => (2, -3),
        289 => (0, -3),
        314 => (1, -3),
        320 => (1, -3),
        327 => (1, -2.5),
        337 => (0, 3),
        342 => (2, 3),
        349 => (4, 0),
        387 => (2, 3),
        388 => (4, 0),
        406 => (4, 0),
        458 => (-4, 0),
        502 => (-2, 3),
        522 => (-3, -2),
        539 => (-4, -2),
        556 => (-1, -3),
        562 => (-3, 2),
        563 => (-4, 1),
        611 => (-2, 3),
        614 => (-2, -3),
        619 => (-1, 3),
        629 => (-4, 1),
        639 => (0, 3),
        676 => (-2, 3),
        682 => (0, 3),
        688 => (-2, -3),
        701 => (3, -2),
        702 => (2, -3),
        755 => (2, -3),
        778 => (0, -3),
        780 => (2, -3),
        785 => (-4, 1),
        813 => (2, -3),
        817 => (-3, -3),
        835 => (2, -3),
        860 => (4, -1),
        861 => (-2, 3),
        886 => (2, -3),
        896 => (2, -3),
        898 => (3, -3),
        899 => (2, -3),
        900 => (4, -1),
        906 => (0, -3)
    )


Zone1 = [101, 226, 166]
Zone2 = [247, 332, 604]
Zone3 = [373, 666, 839]
Zone = vcat(Zone1, Zone2, Zone3)

Name = Dict(
    101 => "Z1.1",
    226 => "Z1.2",
    166 => "Z1.3",
    247 => "Z2.1",
    332 => "Z2.2",
    604 => "Z2.3",
    373 => "Z3.1",
    666 => "Z3.2",
    839 => "Z3.3"
)

offsets1 = Dict(
    25 => (-1, -4),
    32 => (-1, -4),
    59 => (-1, -4),
    101 => (-3, -4),
    127 => (5, -0),
    145 => (5, 0),
    155 => (0, 5),
    166 => (8, 0),
    171 => (0, 5),
    196 => (-5, 0),
    226 => (6, 0),
    247 => (-6, 2),
    263 => (5, 0.0),
    310 => (0, 5),
    325 => (0, 5),
    332 => (0, -6),
    453 => (-5, 0),
    475 => (-5, 0),
    508 => (-5, 0),
    559 => (-5, 1),
    596 => (-4.5, 2),
    604 => (4, 5),
    373 => (-3, -4),
    391 => (-1, -5),
    578 => (-4, 2),
    666 => (0, 7),
    686 => (5, 0),
    691 => (5, 0),
    707 => (5, 0),
    718 => (5, -0.5),
    745 => (5, -0.5),
    794 => (5, -0.5),
    839 => (7, -2)
)

#offsets = Dict(k => (v[1] * 1.25, v[2] * 1.25) for (k, v) in offsets)
# Define the path to your Excel file and sheet name
file_path = "DSSGraph_Output.xlsx"  # Replace with your actual file path
sheet_name = "DSSGraph_Output"      # Replace with your actual sheet name

# Read the data range from the specified sheet
data_range = "A2:E907"  # Update the range according to your data
matrix_data = XLSX.readdata(file_path, sheet_name, data_range)

# Store each column in separate vectors
Bus = matrix_data[:, 1]  # Bus Number
X1_values = matrix_data[:, 2]  # X1
Y1_values = matrix_data[:, 3]  # Y1
X2_values = matrix_data[:, 4]  # X2
Y2_values = matrix_data[:, 5]  # Y2

# Initialize a plot
p = plot(
    xlabel = "X Coordinate",
    ylabel = "Y Coordinate",
    legend = false,
    grid = false,
    framestyle=:none,
    size=(6300, 4000),
    topmargin=2Plots.Measures.cm
)

annotate!(p, (X1_values[2], Y1_values[2]-1.25, text("*", :red, 180)))
annotate!(p, (X1_values[2], Y1_values[2]+5, text("Substation", :red, 110)))
# Add lines to the plot
for i in 2:length(X1_values)
    X1, Y1, X2, Y2 = X1_values[i], Y1_values[i], X2_values[i], Y2_values[i]
    plot!([X1, X2], [Y1, Y2], lw=15, color=:blue, alpha=0.9)
    
end

for j in Zone1
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-1.0, text("*", :red, 150)))
end

for j in Zone2
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-1.0, text("*", :red, 150)))
end

for j in Zone3
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-1.0, text("*", :red, 150)))
end

CFT = 1.2;
for j in Zone
    X2, Y2 = X2_values[j], Y2_values[j]
    X2 = X2 + (offsets1[j][1]*CFT)  # Horizontal offset
    Y2 = Y2 + (offsets1[j][2]*CFT)  # Vertical offset
    annotate!(p, (X2, Y2, text(string(Name[j]), :black, 120)))
end

# Show the plot
display(p)

# Save the plot as an image
savefig(p, "SingleLineDiagramLines.svg")