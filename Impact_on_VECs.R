# Loading required libraries
library(sf)
library(dplyr)
library(openxlsx)

# 1. SETTING WORKING DIRECTORY (update as needed)
setwd("G:/Shared drives/EGIS Active Projects/24-26 Yinta Referral Tools Assessment")

# 2. DEFINING FILE PATHS
house_fp   <- "24-26 Update YLS Scripts/Layers/Refined Data/Boundaries/house_territory_wetsuweten_2012_edit-2021.shp"
yintah_fp  <- "24-26 Update YLS Scripts/Layers/Refined Data/Boundaries/yinta_outline.shp"
vec_fp     <- "24-26 Update YLS Scripts/Layers/Refined Data/Output/combined/berries-comb_class_py_OW_yinta_EPSG3005_2021_07_31.gpkg"
project_fp <- "24-26 Tenas Coal Assessment/Boundaries/22-09_Tena_Coal_Licences_Boundary.shp"

# 3. READING THE DATA
house_territories <- st_read(house_fp, quiet = TRUE)
yintah            <- st_read(yintah_fp, quiet = TRUE)
vec               <- st_read(vec_fp, quiet = TRUE)
project           <- st_read(project_fp, quiet = TRUE)

# Printing column names to verify structure
print("House territories columns:")
print(names(house_territories))
print("VEC columns:")
print(names(vec))

# 4. TRANSFORMING ALL DATA TO CRS EPSG:3005
house_territories <- st_transform(house_territories, 3005)
yintah            <- st_transform(yintah, 3005)
vec               <- st_transform(vec, 3005)
project           <- st_transform(project, 3005)

# 5. SELECTING HOUSE TERRITORIES THAT OVERLAP THE PROJECT BOUNDARY
overlap_matrix <- st_intersects(house_territories, project)
overlapping_indices <- which(lengths(overlap_matrix) > 0)
houses_overlap <- house_territories[overlapping_indices, ]

# Defining the VEC name
vec_name <- "Berries"

# Creating a list to hold results for Excel sheets
sheet_list <- list()

# 6. PROCESSING EACH OVERLAPPING HOUSE TERRITORY
for(i in seq_len(nrow(houses_overlap))) {
  # Extract the current house polygon
  current_house <- houses_overlap[i, ]
  
  # Clipping the project boundary to this house territory
  clipped_project <- st_intersection(project, current_house)
  
  # Skipping if no valid intersection
  if(nrow(clipped_project) == 0) next
  
  # Clipping VEC to the current house territory first
  vec_in_house <- st_intersection(vec, current_house)
  
  # Skipping if no VEC present in this house
  if(nrow(vec_in_house) == 0) next
  
  # Calculating total VEC area in the house
  vec_in_house <- vec_in_house %>%
    mutate(area_km2 = as.numeric(st_area(.)) / 1e6)
  
  total_vec_area <- sum(vec_in_house$area_km2)
  
  # Grouping by classification for full VEC in house
  total_summary <- vec_in_house %>%
    group_by(COMB_CLASS) %>%
    summarise(CLASSIFIED_AREA_KM2 = sum(area_km2)) %>%
    ungroup() %>%
    mutate(PERCENT_OF_TTY = round((CLASSIFIED_AREA_KM2 / total_vec_area) * 100, 1))
  
  # Calculating impacted areas using the clipped project boundary
  impacted_vec <- st_intersection(vec_in_house, clipped_project)
  
  if(nrow(impacted_vec) > 0) {
    impacted_vec <- impacted_vec %>%
      mutate(impact_area_km2 = as.numeric(st_area(.)) / 1e6)
    
    impact_summary <- impacted_vec %>%
      group_by(COMB_CLASS) %>%
      summarise(VEC_AREA_IMPACTED_BY_PROJECT_KM2 = sum(impact_area_km2)) %>%
      ungroup()
  } else {
    impact_summary <- data.frame(
      COMB_CLASS = unique(vec_in_house$COMB_CLASS),
      VEC_AREA_IMPACTED_BY_PROJECT_KM2 = 0
    )
  }
  
  # Removing geometry for joining
  total_summary_df <- st_drop_geometry(total_summary)
  impact_summary_df <- st_drop_geometry(impact_summary)
  
  # Joining summaries and calculate impact percentages
  result_house <- left_join(total_summary_df, impact_summary_df, by = "COMB_CLASS") %>%
    mutate(
      VEC_AREA_IMPACTED_BY_PROJECT_KM2 = coalesce(VEC_AREA_IMPACTED_BY_PROJECT_KM2, 0),
      PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = case_when(
        CLASSIFIED_AREA_KM2 > 0 ~ round((VEC_AREA_IMPACTED_BY_PROJECT_KM2 / CLASSIFIED_AREA_KM2) * 100, 1),
        TRUE ~ 0
      )
    )
  
  # Adding house territory information using correct column names
  house_info <- st_drop_geometry(current_house[1, ])
  result_house <- result_house %>%
    mutate(
      HOUSE_TTY_CODE = house_info$HS_TTY_CD,  # Changed to match your column name
      HOUSE = house_info$HOUSE,
      TTY_CODE = house_info$TTY_CODE,
      TTY_NAME = house_info$TTY_NAME,
      CLAN_WET1 = house_info$CLAN_WET1,
      VEC = vec_name,
      VEC_AREA_KM2 = round(total_vec_area, 2)
    )
  
  # Renaming and reorganize columns
  result_house <- result_house %>%
    rename(COMBINED_CLASSIFICATION = COMB_CLASS) %>%
    select(HOUSE_TTY_CODE, HOUSE, TTY_CODE, TTY_NAME, CLAN_WET1, 
           VEC, VEC_AREA_KM2, CLASSIFIED_AREA_KM2, PERCENT_OF_TTY,
           VEC_AREA_IMPACTED_BY_PROJECT_KM2, PERCENT_VEC_AREA_IMPACTED_BY_PROJECT,
           COMBINED_CLASSIFICATION)
  
  # Formatting percentages
  result_house <- result_house %>%
    mutate(
      PERCENT_OF_TTY = paste0(PERCENT_OF_TTY, "%"),
      PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = paste0(PERCENT_VEC_AREA_IMPACTED_BY_PROJECT, "%")
    )
  
  # Adding to sheet list
  sheet_name <- paste(vec_name, result_house$TTY_CODE[1], sep = "_")
  sheet_list[[sheet_name]] <- result_house
}


# 7. PROCESSING THE YINTAH TERRITORY
vec_yintah <- st_intersection(vec, yintah)

if(nrow(vec_yintah) > 0) {
  vec_yintah <- vec_yintah %>%
    mutate(area_km2 = as.numeric(st_area(.)) / 1e6)
  
  total_yintah_area <- sum(vec_yintah$area_km2)
  
  yintah_summary <- vec_yintah %>%
    group_by(COMB_CLASS) %>%
    summarise(CLASSIFIED_AREA_KM2 = sum(area_km2)) %>%
    ungroup() %>%
    mutate(PERCENT_OF_YINTAH = round((CLASSIFIED_AREA_KM2 / total_yintah_area) * 100, 1))
  
  impacted_yintah <- st_intersection(vec_yintah, project)
  
  if(nrow(impacted_yintah) > 0) {
    impacted_yintah <- impacted_yintah %>%
      mutate(impact_area_km2 = as.numeric(st_area(.)) / 1e6)
    
    impact_summary <- impacted_yintah %>%
      group_by(COMB_CLASS) %>%
      summarise(VEC_AREA_IMPACTED_BY_PROJECT_KM2 = sum(impact_area_km2)) %>%
      ungroup()
  } else {
    impact_summary <- data.frame(
      COMB_CLASS = unique(vec_yintah$COMB_CLASS),
      VEC_AREA_IMPACTED_BY_PROJECT_KM2 = 0
    )
  }
  
  result_yintah <- left_join(
    st_drop_geometry(yintah_summary),
    st_drop_geometry(impact_summary),
    by = "COMB_CLASS"
  ) %>%
    mutate(
      VEC_AREA_IMPACTED_BY_PROJECT_KM2 = coalesce(VEC_AREA_IMPACTED_BY_PROJECT_KM2, 0),
      PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = case_when(
        CLASSIFIED_AREA_KM2 > 0 ~ round((VEC_AREA_IMPACTED_BY_PROJECT_KM2 / CLASSIFIED_AREA_KM2) * 100, 1),
        TRUE ~ 0
      ),
      VEC = vec_name,
      VEC_AREA_KM2 = round(total_yintah_area, 2)
    )
  
  result_yintah <- result_yintah %>%
    rename(COMBINED_CLASSIFICATION = COMB_CLASS) %>%
    select(VEC, VEC_AREA_KM2, CLASSIFIED_AREA_KM2, PERCENT_OF_YINTAH,
           VEC_AREA_IMPACTED_BY_PROJECT_KM2, PERCENT_VEC_AREA_IMPACTED_BY_PROJECT,
           COMBINED_CLASSIFICATION)
  
  result_yintah <- result_yintah %>%
    mutate(
      PERCENT_OF_YINTAH = paste0(PERCENT_OF_YINTAH, "%"),
      PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = paste0(PERCENT_VEC_AREA_IMPACTED_BY_PROJECT, "%")
    )
  
  sheet_list[[paste(vec_name, "Yintah", sep = "_")]] <- result_yintah
}

# 8. WRITING RESULTS TO EXCEL
output_file <- "Testing/VEC_Impact_Results.xlsx"
wb <- createWorkbook()

for(sheet in names(sheet_list)) {
  addWorksheet(wb, sheet)
  writeData(wb, sheet, sheet_list[[sheet]])
}

saveWorkbook(wb, output_file, overwrite = TRUE)
cat("Results saved to", output_file, "\n")
