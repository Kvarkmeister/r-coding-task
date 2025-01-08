library(here)
library(readxl)
library(ggplot2)
library(gridExtra)

import_data <- function(filepath, sheetname, title, start_row, end_row) {
  tryCatch({
    data <- read_excel(filepath, sheet = sheetname, range = cell_limits(c(start_row, 1), c(end_row, NA)), col_names = TRUE)
    colnames(data) <- c("Sample", title, "SD")
    return(data)
  }, error = function(e) {
    print(paste("Error importing data:", e$message))
    return(NULL)
  })
}

create_bar_chart <- function(data, title, y_label, col_name) {
  ggplot(data, aes(x = Sample, y = get(col_name), fill = Sample)) + 
    geom_bar(stat = "identity", position = "dodge") +
    geom_errorbar(aes(ymin = get(col_name) - SD, ymax = get(col_name) + SD), width = 0.2, position = position_dodge(0.9)) + 
    scale_fill_manual(values = c("steelblue", "darkorange", "forestgreen", "darkred")) +  # Custom colors
    labs(title = title, y = y_label, x = "Sample") +
    theme_minimal() +  # Default theme without custom fonts
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          legend.position = "none")  # Remove legend
}

main <- function() {
  excel_file <- here("Exercise Group A.xlsx")
  
  suppressWarnings({
    surface_roughness_data <- import_data(excel_file, "Barchart", "Ra", 3, 7)
    hardness_data <- import_data(excel_file, "Barchart", "BH", 12, 16)
  })
  
  if (!is.null(surface_roughness_data) && nrow(surface_roughness_data) > 0) {
    surface_roughness_chart <- create_bar_chart(surface_roughness_data, "Surface Roughness (μm)", "Ra", "Ra")
  }
  
  if (!is.null(hardness_data) && nrow(hardness_data) > 0) {
    hardness_chart <- create_bar_chart(hardness_data, "Brinell Hardness (BH)", "BH", "BH")
  }
  
  if (!is.null(surface_roughness_chart) && !is.null(hardness_chart)) {
    ggsave("charts/bar-chart.png", plot = grid.arrange(surface_roughness_chart, hardness_chart, ncol = 1), width = 8, height = 12, dpi = 300)
  }
}

main()
