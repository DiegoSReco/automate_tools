## Packages ----
library(pacman)
p_load('tidyverse', 'readxl')

excel_sheet_stacked <- function(file_path, pathstr) {
  #Sheet names
  sheet_names <- excel_sheets(file_path)
  #Keep the sheet with the path
  matching_sheets <- sheet_names[grep(pathstr, sheet_names, ignore.case = TRUE)]
  #loading each sheet in a list
  df_list <- list() 
  for (sheet_name in matching_sheets) {
    df <- read_excel(file_path, sheet = sheet_name)
    df <- df %>% mutate(across(everything(), as.character))
    df_list[[sheet_name]] <- df
  }
  #Stacking Data Frame
  df_stacked <- bind_rows(df_list, .id = "sheet_name")
  return(df_stacked)
}
