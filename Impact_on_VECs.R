# Loading required libraries
library(sf)
library(dplyr)
library(openxlsx)

# 1. SETTING WORKING DIRECTORY (update as needed)
setwd("G:/Shared drives/EGIS Active Projects/24-26 Yinta Referral Tools Assessment")

# 2. DEFINE FILE PATHS
house_fp   <- "24-26 Update YLS Scripts/Layers/Refined Data/Boundaries/house_territory_wetsuweten_2012_edit-2021.shp"
yintah_fp  <- "24-26 Update YLS Scripts/Layers/Refined Data/Boundaries/yinta_outline.shp"
vec_fp     <- "24-26 Update YLS Scripts/Layers/Refined Data/Output/combined/berries-comb_class_py_OW_yinta_EPSG3005_2021_07_31.gpkg"
project_fp <- "24-26 Tenas Coal Assessment/Boundaries/22-09_Tena_Coal_Licences_Boundary.shp"

# 3. READING THE DATA
house_territories <- st_read(house_fp, quiet = TRUE)
yintah            <- st_read(yintah_fp, quiet = TRUE)
vec               <- st_read(vec_fp, quiet = TRUE)
project           <- st_read(project_fp, quiet = TRUE)

# 4. TRANSFORMING ALL DATA TO CRS EPSG:3005
house_territories <- st_transform(house_territories, 3005)
yintah            <- st_transform(yintah, 3005)
vec               <- st_transform(vec, 3005)
project           <- st_transform(project, 3005)

# 5. SELECTING HOUSE TERRITORIES THAT OVERLAP THE PROJECT BOUNDARY
# Using st_intersects ensures we only process houses affected by the project.
overlap_logical <- as.logical(st_intersects(house_territories, project, sparse = FALSE)[,1])
houses_overlap  <- house_territories[overlap_logical, ]

# Defining the VEC name (e.g., "Berries")
vec_name <- "Berries"

# Creating a list to hold results for Excel sheets
sheet_list <- list()

# 6. PROCESSING EACH OVERLAPPING HOUSE TERRITORY
for(i in seq_len(nrow(houses_overlap))) {
  
  # Extracting the current house polygon
  ht <- houses_overlap[i, ]
  
  # Defining the impact zone as the intersection of the house with the project.
  impact_zone <- st_intersection(ht, project)
  # Skip if no valid impact zone (should not happen as we filtered but check for safety)
  if(nrow(impact_zone) == 0 || all(st_is_empty(impact_zone))) next
  
  # --- Total VEC Area within the House Territory ---
  # Clipping VEC to the house territory.
  vec_house <- st_intersection(vec, ht)
  if(nrow(vec_house) == 0) next  # Skip if no VEC present in this house
  
  # Calculating area (in km²) for each polygon within the house territory
  vec_house <- vec_house %>% 
    mutate(area_km2 = as.numeric(st_area(.)) / 1e6)
  
  # Total VEC area inside the house territory
  total_area_house <- sum(vec_house$area_km2)
  
  # Grouping by combined classification for full VEC in the house
  total_summary <- vec_house %>%
    group_by(COMB_CLASS) %>%
    summarise(CLASSIFIED_AREA_KM2 = sum(area_km2)) %>%
    ungroup() %>%
    mutate(PERCENT_OF_TTY = ifelse(total_area_house > 0,
                                   round((CLASSIFIED_AREA_KM2 / total_area_house) * 100, 1),
                                   0))
  
  # --- Impacted VEC Area within the House Territory ---
  impacted_house <- st_intersection(vec_house, impact_zone)
  
  if(nrow(impacted_house) > 0) {
    impacted_house <- impacted_house %>% 
      mutate(impact_area_km2 = as.numeric(st_area(.)) / 1e6)
    
    impact_summary <- impacted_house %>%
      group_by(COMB_CLASS) %>%
      summarise(VEC_AREA_IMPACTED_BY_PROJECT_KM2 = sum(impact_area_km2)) %>%
      ungroup()
  } else {
    # Create a summary with zero impact if none found.
    impact_summary <- data.frame(COMB_CLASS = total_summary$COMB_CLASS,
                                 VEC_AREA_IMPACTED_BY_PROJECT_KM2 = 0)
  }
  
  # Removing spatial geometry from summary tables prior to join
  total_summary_df  <- st_set_geometry(total_summary, NULL)
  impact_summary_df <- st_set_geometry(impact_summary, NULL)
  
  # Merging the full summary with the impact summary by the classification field
  result_ht <- left_join(total_summary_df, impact_summary_df, by = "COMB_CLASS")
  # Replace any NA impacted area with 0.
  result_ht$VEC_AREA_IMPACTED_BY_PROJECT_KM2[is.na(result_ht$VEC_AREA_IMPACTED_BY_PROJECT_KM2)] <- 0
  
  # Computing the percentage of each classification’s area impacted by the project.
  result_ht <- result_ht %>%
    mutate(PERCENT_VEC_AREA_IMPACTED_BY_PROJECT =
             if_else(CLASSIFIED_AREA_KM2 > 0,
                     round((VEC_AREA_IMPACTED_BY_PROJECT_KM2 / CLASSIFIED_AREA_KM2) * 100, 1),
                     0))
  
  # Adding identifying house territory attributes (using the first VEC record's attribute)
  vec_attr <- st_set_geometry(vec_house[1, ], NULL)
  result_ht <- result_ht %>%
    mutate(HOUSE_TTY_CODE = vec_attr$HOUSE_TTY_CODE,
           HOUSE          = vec_attr$HOUSE,
           TTY_CODE       = vec_attr$TTY_CODE,
           TTY_NAME       = vec_attr$TTY_NAME,
           CLAN_WET1      = vec_attr$CLAN_WET1,
           VEC            = vec_name,
           VEC_AREA_KM2   = round(total_area_house, 2)
    ) %>%
    select(HOUSE_TTY_CODE, HOUSE, TTY_CODE, TTY_NAME, CLAN_WET1, VEC, VEC_AREA_KM2,
           CLASSIFIED_AREA_KM2, PERCENT_OF_TTY,
           VEC_AREA_IMPACTED_BY_PROJECT_KM2, PERCENT_VEC_AREA_IMPACTED_BY_PROJECT,
           COMB_CLASS)
  
  # Renaming COMB_CLASS to COMBINED_CLASSIFICATION to match sample output
  result_ht <- result_ht %>% rename(COMBINED_CLASSIFICATION = COMB_CLASS)
  
  # Formating the percentage fields with a "%" symbol for display clarity.
  result_ht <- result_ht %>%
    mutate(PERCENT_OF_TTY = paste0(PERCENT_OF_TTY, "%"),
           PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = paste0(PERCENT_VEC_AREA_IMPACTED_BY_PROJECT, "%"))
  
  # Defining the sheet name (example: "Berries_W01A")
  sheet_name <- paste(vec_name, result_ht$TTY_CODE[1], sep = "_")
  
  # Saving the result for this house territory only if there is an impact zone present.
  sheet_list[[sheet_name]] <- result_ht
}

# 7. PROCESSING THE OVERALL YINTAH TERRITORY
vec_yintah <- st_intersection(vec, yintah)
if(nrow(vec_yintah) > 0) {
  vec_yintah <- vec_yintah %>%
    mutate(area_km2 = as.numeric(st_area(.)) / 1e6)
  total_area_yintah <- sum(vec_yintah$area_km2)
  
  summary_yintah <- vec_yintah %>%
    group_by(COMB_CLASS) %>%
    summarise(CLASSIFIED_AREA_KM2 = sum(area_km2)) %>%
    ungroup() %>%
    mutate(PERCENT_OF_YINTAH = ifelse(total_area_yintah > 0,
                                      round((CLASSIFIED_AREA_KM2 / total_area_yintah) * 100, 1),
                                      0))
  
  # Clipping the VEC with the entire project boundary for Yintah.
  impacted_yintah <- st_intersection(vec_yintah, project)
  if(nrow(impacted_yintah) > 0) {
    impacted_yintah <- impacted_yintah %>%
      mutate(impact_area_km2 = as.numeric(st_area(.)) / 1e6)
    impact_yintah_summary <- impacted_yintah %>%
      group_by(COMB_CLASS) %>%
      summarise(VEC_AREA_IMPACTED_BY_PROJECT_KM2 = sum(impact_area_km2)) %>%
      ungroup()
  } else {
    impact_yintah_summary <- data.frame(COMB_CLASS = summary_yintah$COMB_CLASS,
                                        VEC_AREA_IMPACTED_BY_PROJECT_KM2 = 0)
  }
  
  # Remove spatial components for joining
  summary_yintah_df    <- st_set_geometry(summary_yintah, NULL)
  impact_yintah_df     <- st_set_geometry(impact_yintah_summary, NULL)
  
  result_yintah <- left_join(summary_yintah_df, impact_yintah_df, by = "COMB_CLASS")
  result_yintah$VEC_AREA_IMPACTED_BY_PROJECT_KM2[is.na(result_yintah$VEC_AREA_IMPACTED_BY_PROJECT_KM2)] <- 0
  
  result_yintah <- result_yintah %>%
    mutate(PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = if_else(
      CLASSIFIED_AREA_KM2 > 0,
      round((VEC_AREA_IMPACTED_BY_PROJECT_KM2 / CLASSIFIED_AREA_KM2) * 100, 1),
      0
    ),
    VEC = vec_name,
    VEC_AREA_KM2 = round(total_area_yintah, 2)) %>%
    select(VEC, VEC_AREA_KM2, CLASSIFIED_AREA_KM2, PERCENT_OF_YINTAH,
           VEC_AREA_IMPACTED_BY_PROJECT_KM2, PERCENT_VEC_AREA_IMPACTED_BY_PROJECT, COMB_CLASS)
  
  result_yintah <- result_yintah %>% rename(COMBINED_CLASSIFICATION = COMB_CLASS)
  
  result_yintah <- result_yintah %>%
    mutate(PERCENT_OF_YINTAH = paste0(PERCENT_OF_YINTAH, "%"),
           PERCENT_VEC_AREA_IMPACTED_BY_PROJECT = paste0(PERCENT_VEC_AREA_IMPACTED_BY_PROJECT, "%"))
  
  # Saving the Yintah summary sheet (example: "Berries_Yintah")
  sheet_list[[paste(vec_name, "Yintah", sep = "_")]] <- result_yintah
}

# 8. WRITING ALL RESULTS TO A SINGLE EXCEL FILE WITH MULTIPLE SHEETS
output_file <- "Testing/VEC_Impact_Results.xlsx"
wb <- createWorkbook()

for(sheet in names(sheet_list)) {
  addWorksheet(wb, sheet)
  writeData(wb, sheet, sheet_list[[sheet]])
}

saveWorkbook(wb, output_file, overwrite = TRUE)
cat("Results saved to", output_file, "\n")
