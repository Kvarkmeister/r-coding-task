# Load necessary libraries
library(readxl)
library(dplyr)
library(car)
library(multcomp)
library(openxlsx)
library(here)

# Create the 'excels' directory if it doesn't exist
if (!dir.exists("excels")) {
  dir.create("excels")
}

# Load the data from the Excel file
file_path <- "Exercise Group A.xlsx"
sheet_name <- "R statistics"

# Read the data for Surface roughness (Ra)
data_ra <- read_excel(file_path, sheet = sheet_name, range = "A4:B44", col_names = c("Sample", "Ra"))

# Read the data for Brinell hardness (BH)
data_bh <- read_excel(file_path, sheet = sheet_name, range = "E4:F28", col_names = c("Sample", "BH"))

# Convert Ra and BH columns to numeric
data_ra$Ra <- as.numeric(data_ra$Ra)
data_bh$BH <- as.numeric(data_bh$BH)

# Debugging: Check for any NA/NaN/Inf values in the data
na_ra <- data_ra %>% filter(is.na(Ra) | is.infinite(Ra))
na_bh <- data_bh %>% filter(is.na(BH) | is.infinite(BH))

print("Rows with NA/NaN/Inf values in Surface roughness (Ra) data:")
print(na_ra)

print("Rows with NA/NaN/Inf values in Brinell hardness (BH) data:")
print(na_bh)

# Clean the data by removing rows with NA/NaN/Inf values
data_ra <- data_ra %>% filter(!is.na(Ra) & !is.infinite(Ra))
data_bh <- data_bh %>% filter(!is.na(BH) & !is.infinite(BH))

# Check for any remaining NA/NaN/Inf values
print(sum(is.na(data_ra$Ra)))
print(sum(is.na(data_bh$BH)))

# Debugging: Check the first few rows of the data
print("First few rows of Surface roughness (Ra) data:")
print(head(data_ra))

print("First few rows of Brinell hardness (BH) data:")
print(head(data_bh))

# Perform ANOVA analysis for Surface roughness (Ra)
anova_ra <- aov(Ra ~ Sample, data = data_ra)
summary_anova_ra <- summary(anova_ra)
anova_table_ra <- as.data.frame(summary(anova_ra)[[1]])
print("ANOVA results for Surface roughness (Ra):")
print(anova_table_ra)

# Perform ANOVA analysis for Brinell hardness (BH)
anova_bh <- aov(BH ~ Sample, data = data_bh)
summary_anova_bh <- summary(anova_bh)
anova_table_bh <- as.data.frame(summary(anova_bh)[[1]])
print("ANOVA results for Brinell hardness (BH):")
print(anova_table_bh)

# Perform Tukey HSD test with confidence level of 95% for Surface roughness (Ra)
tukey_hsd_ra <- TukeyHSD(anova_ra, conf.level = 0.95)
print("Tukey HSD results for Surface roughness (Ra):")
print(tukey_hsd_ra)

# Perform Tukey HSD test with confidence level of 95% for Brinell hardness (BH)
tukey_hsd_bh <- TukeyHSD(anova_bh, conf.level = 0.95)
print("Tukey HSD results for Brinell hardness (BH):")
print(tukey_hsd_bh)

# Create groupings of Tukey HSD results for Surface roughness (Ra)
groupings_ra <- as.data.frame(tukey_hsd_ra$Sample)
groupings_ra$Comparison <- rownames(groupings_ra)

# Create groupings of Tukey HSD results for Brinell hardness (BH)
groupings_bh <- as.data.frame(tukey_hsd_bh$Sample)
groupings_bh$Comparison <- rownames(groupings_bh)

# Specify the full path for the output files using here()
output_file_path_anova <- here("excels/anova-analysis.xlsx")
output_file_path_tukey <- here("excels/TUKEYHSD-analysis.xlsx")

# Export the ANOVA results to an Excel file
write.xlsx(list("ANOVA_Surface_Roughness_Ra" = anova_table_ra,
                "ANOVA_Brinell_Hardness_BH" = anova_table_bh), 
           file = output_file_path_anova)

# Export the Tukey HSD results to an Excel file
write.xlsx(list("TukeyHSD_Surface_Roughness_Ra" = groupings_ra,
                "TukeyHSD_Brinell_Hardness_BH" = groupings_bh), 
           file = output_file_path_tukey)

print(paste("ANOVA results have been exported to", output_file_path_anova))
print(paste("Tukey HSD results have been exported to", output_file_path_tukey))