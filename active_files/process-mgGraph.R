if (!require("pacman")) install.packages("pacman") 
pacman::p_load(
  data.table, tidyverse, janitor, lubridate, # Visualization
  ggplot2, ggthemes, plotly, leaflet, esquisse, # Reporting
  DT, gdata, kable, kableExtra, tufte, # Interactive tools
  shiny )

# Define a reusable function for importing and cleaning data

drop_columns_by_values_and_patterns <- function(file_path, match_strings = c("Microsoft", "System")) {
  # Read CSV
  dt <- fread(file_path, na.strings = c("", "NA"))
  
  # Drop columns where all values are NA or empty
  non_empty_cols <- sapply(dt, function(col) any(!is.na(col) & col != ""))
  
  # Drop columns where any value matches any of the test strings
  match_cols <- sapply(dt, function(col) {
    any(sapply(match_strings, function(pattern) {
      any(grepl(pattern, as.character(col), ignore.case = TRUE))
    }))
  })
  
  # Combine filters: keep columns that are non-empty and do not match any pattern
  cols_to_keep <- names(dt)[non_empty_cols & !match_cols]
  
  # Return cleaned data.table
  return(dt[, ..cols_to_keep])
}

#Set working directory (consider replacing with here::here() for portability)
setwd("C:/scripts/mgReports")

# function usage
 user_data <- drop_columns_by_values_and_patterns("mgUsers.csv", c("Microsoft", "System"))
 group_data <- drop_columns_by_values_and_patterns("mgGroups.csv", c("Microsoft", "System"))
 u_group_data <- drop_columns_by_values_and_patterns("mgUnifiedGroups.csv", c("Microsoft", "System"))


