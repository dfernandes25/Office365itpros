




if (!require("pacman")) install.packages("pacman") 
pacman::p_load(
  data.table, tidyverse, janitor, lubridate, # Visualization
  ggplot2, ggthemes, jsonlite, plotly, leaflet, esquisse, # Reporting
  DT, gdata, kable, kableExtra, tufte, # Interactive tools
  shiny )

parse_json_directory <- function(directory_path) {
  # Find all .json files recursively
  json_files <- list.files(path = directory_path, pattern = "\\.json$", full.names = TRUE, recursive = TRUE)
  
  # Initialize a list to hold data frames
  df_list <- list()
  
  # Loop through each file and read it
  for (file in json_files) {
    message("Reading: ", file)
    json_data <- fromJSON(file, flatten = TRUE)
    
    # Convert to data.frame and then to data.table
    df <- as.data.table(json_data)
    
    # Append to list
    df_list[[length(df_list) + 1]] <- df
  }
  
  # Combine all data.tables into one
  combined_dt <- rbindlist(df_list, fill = TRUE)
  
  return(combined_dt)
}


combined_logs <- parse_json_directory("C:/scripts/user_logins")
combined_logs$createdDateTime <- format(ymd_hms(combined_logs$createdDateTime), "%m/%d/%Y")

 tmp1 <- combined_logs %>%
   group_by(appDisplayName, userPrincipalName) %>%
  summarise(
     count = n()
  )