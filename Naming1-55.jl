selected_indices = [70, 34, 47, 83, 73, 74, 178, 208, 264, 225, 248, 249, 327, 320, 314, 276,
                    289, 387, 342, 388, 349, 337, 406, 629, 563, 502, 611, 562, 682, 676, 619, 639,
                    458, 556, 539, 522, 785, 614, 688, 817, 860, 861, 778, 701, 702, 900, 896, 755, 
                    813, 780, 835, 898, 906, 886, 899]

rooftop_set = [629, 886, 906, 337, 556, 458, 34, 248, 320, 249, 289, 682, 701, 896]

index_dict = Dict(70 => 1, 34 => 2, 47 => 3, 83 => 4, 73 => "5 & 6", 74 => " ", 178 => 7, 208 => 8, 264 => 9, 225 => 10, 
                    248 => "11 & 12", 249 => " ", 327 => 17, 320 => 16, 314 => 15, 276 => 14, 289 => 13, 387 => 18, 342 => 19, 
                    388 => 20, 349 => 21, 337 => 22, 406 => 23, 629 => 24, 563 => 25, 502 => 26, 611 => 27, 562 => 28, 
                    682 => 29, 676 => 30, 619 => 31, 639 => 32, 458 => 33, 556 => 34, 539 => 35, 522 => 36, 785 => 37, 
                    614 => 38, 688 => 39, 817 => 40, 860 => 41, 861 => 42, 778 => 43, 701 => 44, 702 => 45, 900 => 46, 
                    896 => 47, 755 => 48, 813 => 49, 780 => 50, 835 => 51, 898 => 52, 906 => 53, 886 => 54, 899 => 55)
     
                    
using XLSX
using Plots

ENV["GKSwstype"] = "pdf"
ENV["GKS_FILEPATH"] = "SingleLineDiagram2.pdf"
ENV["GKSwstype_DPI"] = "1200"


    offsets = Dict(
        34 => (3, 1),
        47 => (0, 5),
        70 => (2, 3),
        73 => (0, -4),
        74 => (0, -4),
        83 => (3, 0),
        178 => (-1, 4),
        208 => (-4, 1),
        225 => (-4, 2),
        248 => (0, 5),
        249 => (0, 5),
        264 => (-3, 3),
        276 => (2, -5),
        289 => (0, -5),
        314 => (0, -5),
        320 => (2, -4),
        327 => (2, -4.5),
        337 => (3.5, 3.5),
        342 => (2, 5),
        349 => (4, 2),
        387 => (0, 5),
        388 => (4, 3),
        406 => (4, -1),
        458 => (-4, 0),
        502 => (-3, 4),
        522 => (-3, -4),
        539 => (-4, -2),
        556 => (-4, -3),
        562 => (-3, 4),
        563 => (-3, 4),
        611 => (-2, 5),
        614 => (-4, -3),
        619 => (-1, 5),
        629 => (-4, 1),
        639 => (0, 5),
        676 => (-2, 5),
        682 => (0, 5),
        688 => (-3, -4),
        701 => (3, -3),
        702 => (2, -4),
        755 => (3, -4),
        778 => (0, -4),
        780 => (2, -4),
        785 => (-4, 1),
        813 => (0, -4),
        817 => (-2, -4),
        835 => (2, -4),
        860 => (3, -2),
        861 => (-2, 5),
        886 => (2, -4),
        896 => (2, -4),
        898 => (3, -4),
        899 => (2, -4),
        900 => (4, -2),
        906 => (0, -4)
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
    dpi = 300,
    legend = false,
    grid = false,
    framestyle=:none,
    size=(6300, 4000),
    top_margin=2Plots.Measures.cm,
    bottom_margin=1Plots.Measures.cm,
    right_margin=1Plots.Measures.cm
)

annotate!(p, (X1_values[2], Y1_values[2]-1.25, text("*", :red, 180)))
annotate!(p, (X1_values[2], Y1_values[2]+5, text("Substation", :red, 90)))
# Add lines to the plot
for i in 2:length(X1_values)
    X1, Y1, X2, Y2 = X1_values[i], Y1_values[i], X2_values[i], Y2_values[i]
    plot!([X1, X2], [Y1, Y2], lw=18, color=:blue, alpha=1.0)
    
end

for j in setdiff(selected_indices, rooftop_set)
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-0.5, text("*", :black, 120)))
end

for j in rooftop_set
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-0.5, text("*", :red, 140)))
end

for j in [786, 617, 881]
    X2, Y2 = X2_values[j], Y2_values[j]
    annotate!(p, (X2, Y2-1, text("*", :red, 160)))
end

#
for j in selected_indices
    X2, Y2 = X2_values[j], Y2_values[j]
    X2 = X2 + 1 * offsets[j][1]  # Horizontal offset
    Y2 = Y2 + 1 * offsets[j][2]  # Vertical offset
    annotate!(p, (X2, Y2, text(string(index_dict[j]), :black, 100, alpha=1)))
end
#

X2, Y2 = X2_values[786], Y2_values[786]
X2 = X2 + 0  # Horizontal offset
Y2 = Y2 + (-4)  # Vertical offset
annotate!(p, (X2, Y2, text("DER2", :red, 100, alpha=1)))

X2, Y2 = X2_values[617], Y2_values[617]
X2 = X2 + 2  # Horizontal offset
Y2 = Y2 + 5  # Vertical offset
annotate!(p, (X2, Y2, text("DER1", :red, 100, alpha=1)))

X2, Y2 = X2_values[881], Y2_values[881]
X2 = X2 + 0  # Horizontal offset
Y2 = Y2 + (-5)  # Vertical offset
annotate!(p, (X2, Y2, text("DER3", :red, 100, alpha=1)))

# Show the plot
display(p)

# Save the plot as an image
savefig(p, "SingleLineDiagram3.pdf")
savefig(p, "SingleLineDiagram3.png")
savefig(p, "SingleLineDiagram3.svg")
