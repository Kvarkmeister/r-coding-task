library(here)
library(readxl)
library(ggplot2)
library(gridExtra)  # For arranging multiple plots

import_data <- function(filepath, sheetname, title, header_row, start_row, end_row) {
  tryCatch({
    # Import the data with specific header and row range
    data <- read_excel(filepath, sheet = sheetname, range = cell_limits(c(start_row, 1), c(end_row, NA)), col_names = TRUE)
    
    # Print the first few rows to inspect the data
    print(head(data)) 
    
    # Set column names based on the title for correct labeling
    colnames(data) <- c("Sample", title, "SD")
    
    # Return the data as is (no cleaning or conversion)
    return(data)
  }, error = function(e) {
    print(paste("Error importing data:", e$message))
    return(NULL)
  })
}

main <- function() {
  excel_file <- here("Exercise Group A.xlsx")
  print(paste("Resolved file path:", excel_file))
  
  # Suppress warnings while importing data
  suppressWarnings({
    # Import data for Surface Roughness (Rows 3-7)
    surface_roughness_data <- import_data(excel_file, "Barchart", "Ra", header_row = 3, start_row = 3, end_row = 7)
    
    # Adjusted row range for Brinell Hardness (now using the correct header row)
    hardness_data <- import_data(excel_file, "Barchart", "BH", header_row = 11, start_row = 12, end_row = 16)
  })
  
  # Ensure we have valid data for both charts before creating them
  if (!is.null(surface_roughness_data) && nrow(surface_roughness_data) > 0) {
    # Create the Surface Roughness chart
    surface_roughness_chart <- create_bar_chart(surface_roughness_data, "Ra", "Surface Roughness (μm)", "Ra")
  }
  
  if (!is.null(hardness_data) && nrow(hardness_data) > 0) {
    # Create the Hardness chart
    hardness_chart <- create_bar_chart(hardness_data, "BH", "Brinell Hardness (HB)", "BH")
  }
  
  # Combine the two charts into one PNG file vertically
  if (!is.null(surface_roughness_chart) && !is.null(hardness_chart)) {
    # Save both charts to a single image vertically
    ggsave("charts/bar-chart.png", 
           plot = grid.arrange(surface_roughness_chart, hardness_chart, ncol = 1), 
           width = 8, height = 12, dpi = 300)
  }
}

main()
