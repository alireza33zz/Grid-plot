using XLSX
using Plots

ENV["GKSwstype"] = "png"
ENV["GKS_FILEPATH"] = "SingleLineDiagram.png"
ENV["GKSwstype_DPI"] = "300"

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
    title = "Single-Line Diagram of LV_Eu_test",
    legend = false,
    grid = false,
    framestyle=:none,
    size=(6000, 4000),
    titlefontsize=80
)

annotate!(p, (X1_values[2], Y1_values[2]-1.0, text("*", :red, 120)))
annotate!(p, (X1_values[2], Y1_values[2]+5, text("Substation", :red, 70)))
# Add lines to the plot
for i in 2:length(X1_values)
    X1, Y1, X2, Y2 = X1_values[i], Y1_values[i], X2_values[i], Y2_values[i]
    plot!([X1, X2], [Y1, Y2], lw=10, color=:blue, alpha=0.9)
    
end

for j in selected_indices
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-0.5, text("*", :red, 60)))
end

for j in selected_indices
    X2, Y2 = X2_values[j], Y2_values[j]
    X2 = X2 + offsets[j][1]  # Horizontal offset
    Y2 = Y2 + offsets[j][2]  # Vertical offset
    annotate!(p, (X2, Y2, text(string(Bus[j]), :black, 75)))
end


# Show the plot
display(p)

# Save the plot as an image
savefig(p, "SingleLineDiagram.png")