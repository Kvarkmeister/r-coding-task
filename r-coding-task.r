library(here)
library(readxl)
library(ggplot2)
library(gridExtra)

# Function to import data
import_data <- function(filepath, sheetname, start_row, end_row) {
  tryCatch({
    data <- read_excel(filepath, sheet = sheetname, range = cell_limits(c(start_row, 1), c(end_row, 3)), col_names = TRUE)
    return(data)
  }, error = function(e) {
    print(paste("Error importing data:", e$message))
    return(NULL)
  })
}

# Function to create bar charts
create_bar_chart <- function(data, title, y_label, col_name) {
  ggplot(data, aes(x = Sample, y = get(col_name), fill = Sample)) + 
    geom_bar(stat = "identity", position = "dodge") +
    geom_errorbar(aes(ymin = get(col_name) - SD, ymax = get(col_name) + SD), width = 0.2, position = position_dodge(0.9)) + 
    scale_fill_manual(values = c("steelblue", "darkorange", "forestgreen", "darkred")) +  # Custom colors
    labs(title = title, y = y_label, x = "Sample") +
    theme_bw() +  # Set a white background
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          legend.position = "none",  # Remove legend
          plot.title = element_text(hjust = 0.5, size = 16))  # Center plot title
}

# Function to create line plot
create_line_plot <- function(data, title) {
  ggplot(data, aes(x = Ra, y = BH)) + 
    geom_line(aes(color = Sample), linewidth = 1) +  # Replace size with linewidth
    geom_point(aes(color = Sample), size = 3) +
    scale_color_manual(values = c("steelblue", "darkorange", "forestgreen", "darkred")) + 
    labs(title = title, x = "Surface Roughness (Ra)", y = "Brinell Hardness (BH)") +
    theme_bw() +  # Set a white background
    theme(legend.position = "right",  # Move legend to the right
          legend.text = element_text(size = 12),  # Increase legend text size
          legend.title = element_text(size = 14),  # Increase legend title size
          plot.title = element_text(hjust = 0.5, size = 16))  # Center plot title
}

# Main function
main <- function() {
  excel_file <- here("Exercise Group A.xlsx")
  
  suppressWarnings({
    # Importing data from the "Barchart" sheet
    surface_roughness_data <- import_data(excel_file, "Barchart", 3, 7)
    hardness_data <- import_data(excel_file, "Barchart", 12, 16)
    
    # Import line plot data with correct range for rows 4-27 and columns A, B, C
    line_plot_data <- read_excel(excel_file, sheet = "Line plot", range = cell_limits(c(4, 1), c(27, 3)), col_names = TRUE)
    colnames(line_plot_data) <- c("Sample", "Ra", "BH")
    
    # Print the structure of surface_roughness_data, hardness_data, and line_plot_data to verify their content
    print("Surface Roughness Data:")
    print(head(surface_roughness_data))  # Debug: check surface roughness data
    
    print("Hardness Data:")
    print(head(hardness_data))  # Debug: check hardness data
    
    print("Line Plot Data:")
    print(head(line_plot_data))  # Debug: check line plot data
  })
  
  # Create bar chart for surface roughness and hardness
  if (!is.null(surface_roughness_data) && nrow(surface_roughness_data) > 0) {
    surface_roughness_chart <- create_bar_chart(surface_roughness_data, "Surface Roughness (μm)", "Ra", "Ra")
  }
  
  if (!is.null(hardness_data) && nrow(hardness_data) > 0) {
    hardness_chart <- create_bar_chart(hardness_data, "Brinell Hardness (BH)", "BH", "BH")
  }
  
  # Create line plot for Ra vs BH
  if (!is.null(line_plot_data) && nrow(line_plot_data) > 0) {
    line_plot <- create_line_plot(line_plot_data, "Surface Roughness vs Brinell Hardness")
    ggsave("charts/line-plot.png", plot = line_plot, width = 8, height = 6, dpi = 600)  # Save line plot with high DPI
  }
  
  # Save bar charts as a combined image with new file name
  if (!is.null(surface_roughness_chart) && !is.null(hardness_chart)) {
    ggsave("charts/bar-chart.png", 
           plot = grid.arrange(surface_roughness_chart, hardness_chart, ncol = 1), 
           width = 8, height = 12, dpi = 600)  # Save bar charts with high DPI
  }
}

# Run the main function
main()
