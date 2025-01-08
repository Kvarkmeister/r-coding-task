import_data <- function(filepath, sheetname, title) {
  tryCatch({
    # Import the data and skip the first non-relevant rows (titles)
    data <- read_excel(filepath, sheet = sheetname, skip = 2)  # Skip the first two rows to avoid titles
    
    # Remove rows that contain titles or non-numeric values
    data <- data[!is.na(data$Sample) & !data$Sample %in% c("Surface roughness average values", "Brinell hardness average values"), ]
    
    # Ensure that 'Ra'/'BH' and 'SD' columns are numeric, handling possible non-numeric values
    data[[title]] <- as.numeric(data[[title]])
    data$SD <- as.numeric(data$SD)
    
    # Remove any rows that contain NA after conversion
    data <- na.omit(data)
    
    # Explicitly set the column names
    colnames(data) <- c("Sample", title, "SD")
    
    return(data)
  }, error = function(e) {
    print(paste("Error importing data:", e$message))
    return(NULL)
  })
}

create_bar_chart <- function(data, y_variable, title, y_label) {
  if (is.null(data)) return(NULL)
  tryCatch({
    chart <- ggplot(data, aes(x = Sample, y = .data[[y_variable]], fill = Sample)) +
      geom_bar(stat = "identity", position = "dodge") +
      geom_errorbar(aes(ymin = .data[[y_variable]] - SD, ymax = .data[[y_variable]] + SD), 
                    position = position_dodge(width = 0.9), width = 0.2) +
      labs(title = title, x = "Sample", y = y_label, fill = "Sample") +
      theme(text = element_text(family = "Times New Roman"))
    return(chart)
  }, error = function(e) {
    print(paste("Error creating chart:", e$message))
    return(NULL)
  })
}

main <- function() {
  excel_file <- here("Exercise Group A.xlsx")
  print(paste("Resolved file path:", excel_file))
  
  # Import data for both Surface Roughness and Brinell Hardness
  surface_roughness_data <- import_data(excel_file, "Barchart", "Ra")
  hardness_data <- import_data(excel_file, "Barchart", "BH")
  
  # Check for warnings after importing data
  warnings()
  
  # Create the Surface Roughness chart
  surface_roughness_chart <- create_bar_chart(surface_roughness_data, "Ra", "Surface Roughness (μm)", "Ra")
  
  # Create the Hardness chart
  hardness_chart <- create_bar_chart(hardness_data, "BH", "Brinell Hardness (HB)", "BH")
  
  # Save the Surface Roughness chart
  if (!is.null(surface_roughness_chart)) {
    print(surface_roughness_chart)
    ggsave("surface_roughness_chart.png", plot = surface_roughness_chart, width = 8, height = 6, dpi = 300)
  }
  
  # Save the Hardness chart
  if (!is.null(hardness_chart)) {
    print(hardness_chart)
    ggsave("hardness_chart.png", plot = hardness_chart, width = 8, height = 6, dpi = 300)
  }
}

main()
