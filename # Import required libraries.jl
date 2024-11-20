using XLSX
using Plots

selected_indices = [34, 47, 70, 73, 74, 83, 178, 208, 225, 248, 249, 264, 276, 289, 314, 320,
    327, 337, 342, 349, 387, 388, 406, 458, 502, 522, 539, 556, 562, 563, 611, 614, 619, 
    629, 639, 676, 682, 688, 701, 702, 755, 778, 780, 785, 813, 817, 835, 860, 861, 886,
    896, 898, 899, 900, 906]

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
    title = "Single-Line Diagram from Excel",
    legend = false,
    grid = false
)

# Add lines to the plot
for i in 2:length(X1_values)
    X1, Y1, X2, Y2 = X1_values[i], Y1_values[i], X2_values[i], Y2_values[i]
    plot!([X1, X2], [Y1, Y2], lw=2, color=:blue, label="")
end

for j in selected_indices
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2, text(string(Bus[j]), :black, 8)))
end


# Show the plot
display(p)

# Save the plot as an image
savefig(p, "SingleLineDiagram.png")
