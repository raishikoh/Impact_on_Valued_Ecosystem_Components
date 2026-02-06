# Impact Assessment of Infrastructure Projects on Valued Ecosystem Components

This repository contains an R script for assessing project impacts on multiple Valued Ecosystem Components (VECs) within Yintah Territory. The script performs comprehensive spatial overlay analysis, calculates impact metrics by classification category, and generates detailed Excel reports for environmental assessment and consultation processes.

## Installation Guidelines

**Required R Packages:**<br/>
```r
install.packages(c("sf", "dplyr", "openxlsx"))
```

**Setup Instructions:**<br/>
1. Clone or download this repository to your local machine<br/>
2. Install the required R packages using the command above<br/>
3. Set your working directory to the project folder containing spatial data<br/>
4. Update file paths in the configuration section to match your data structure<br/>

## Script Overview

### VEC Impact Assessment Script

This script performs multi-layered spatial impact assessment across multiple VEC datasets and territories, including:

- **Spatial Overlay Analysis**: Intersects project boundaries with VEC layers and territorial boundaries
- **Multi-Territory Processing**: Analyzes impacts across Yintah and individual house territories
- **Classification-Based Metrics**: Calculates impacts by disturbance and designation status
- **Area Calculations**: Computes precise area measurements in square kilometers
- **Percentage Impact Analysis**: Determines proportion of VEC areas affected by project
- **Comprehensive Excel Reporting**: Generates multi-sheet workbooks with summary tables

**What the Script Can Analyze:**
- Total VEC area within Yintah and house territories
- Classified VEC areas (Undisturbed/Disturbed × Designated/Undesignated)
- Project footprint overlap with each VEC classification
- House territories affected by the project boundary
- Percentage of each VEC classification impacted by project
- Spatial distribution of impacts across multiple territories

**What the Script Can Calculate:**
- **Area Metrics**: Total VEC area, classified area, and project impact area (all in km²)
- **Percentage Metrics**: Percent of territory occupied by VEC classifications, percent of VEC impacted by project
- **Summary Statistics**: Overall VEC impacts, cross-territory comparisons, wide-format summary tables
- **Territorial Breakdowns**: House territory codes, names, and territory-specific impact metrics

**Key Features:**
- Automated multi-VEC processing workflow
- Robust spatial intersection operations
- Excel sheet name sanitization for compatibility
- Numeric formatting with appropriate decimal precision
- Hierarchical organization of results by VEC and territory
- Progress monitoring via console output

## Usage Instructions

### Basic Workflow

1. **Configure Working Directory:**
   ```r
   setwd("path/to/your/project/folder")
   ```

2. **Update File Paths:**
   - Update the `house_fp`, `yintah_fp`, and `project_fp` variables to point to your boundary files
   - Configure the `vec_info` list to include all VEC layers with names and file paths
   - Set the `output_file` path for the Excel workbook

3. **Run the Script:**
   ```r
   source("Impact_on_VECs.R")
   ```

4. **Review Output:**
   - Open the generated Excel workbook
   - Review the Summary sheet for consolidated results
   - Examine individual VEC sheets for detailed metrics

### Configuration Options

**Key Variables:**
- `house_fp`: Path to house territory polygons with attributes (TTY_CODE, TTY_NAME, HOUSE)
- `yintah_fp`: Path to Yintah (traditional territory) boundary polygon
- `project_fp`: Path to project assessment boundary polygon
- `vec_info`: List of VEC layers with `name` and `path` elements
- `output_file`: Path for saving the Excel workbook

**VEC Configuration Example:**
```r
vec_info <- list(
  list(name = "Berries", path = "path/to/berries_vec.gpkg"),
  list(name = "Cedar", path = "path/to/cedar_vec.gpkg"),
  list(name = "Moose", path = "path/to/moose_vec.gpkg")
)
```

## Additional Sections

### Features
- **Multi-VEC Processing**: Analyze multiple ecosystem components in a single workflow
- **Territorial Hierarchy**: Assess impacts at both Yintah and house territory levels
- **Classification Analysis**: Break down impacts by disturbance and designation status
- **Automated Excel Generation**: Create professionally formatted multi-sheet workbooks
- **Spatial Precision**: Accurate area calculations using BC Albers projection (EPSG:3005)
- **Flexible Configuration**: Easy customization of VEC layers and file paths
- **Robust Error Handling**: Continues processing even if individual VECs have issues

### Special Instructions
- **COMB_CLASS Requirement**: All VEC layers must have a `COMB_CLASS` field with the four standard classification categories (Undisturbed/Disturbed × Designated/Undesignated)
- **Coordinate Systems**: Input data can be in any CRS - automatic transformation to EPSG:3005 ensures consistency
- **Excel Sheet Names**: Sheet names are automatically sanitized to remove special characters and truncate to 31 characters (Excel limit)
- **House Territory Selection**: Only house territories that spatially overlap the project boundary are included in the analysis
- **Large Datasets**: For projects with many house territories or large VEC layers, processing time may be several minutes
- **Backup Recommendation**: Always review input data quality before running analysis

### VEC Configuration Format
The `vec_info` list defines which VEC layers to analyze:
```r
vec_info <- list(
  list(
    name = "Berries",  # Display name for reports and sheet names
    path = "path/to/berries_vec.gpkg"  # Full path to spatial layer
  ),
  list(
    name = "Cedar",
    path = "path/to/cedar_vec.gpkg"
  )
  # Add additional VECs as needed
)
```

**Supported VEC Types:**
- Berries
- Cedar
- Dry Habitat Species
- Game Birds and Furbearers
- Mountain Goat (Goat)
- Medicinal Plants
- Moose
- Uncommon/Rare Ecosystems

### Output Structure

**Excel Workbook Organization:**

The script generates a single Excel workbook with multiple sheets:

1. **Summary Sheet** - Contains three tables:
   - VEC Impact Summary by Classification (wide-format)
   - House Territories list (TTY_CODE and TTY_NAME)
   - Overall VEC Impacts (total percentage by VEC)

2. **Yintah Sheets** - One sheet per VEC showing territory-wide impacts
   - Naming: `{VEC_Name}_Yintah`
   - Metrics: Total area, classified area, percent of Yintah, impact area, percent impacted

3. **House Territory Sheets** - One sheet per VEC × House combination
   - Naming: `{VEC_Name}_{TTY_CODE}`
   - Metrics: Territory identifiers, VEC areas, impacts by classification

**Classification Categories:**
1. Undisturbed, Designated
2. Undisturbed, Undesignated
3. Disturbed, Designated
4. Disturbed, Undesignated

## Project Information
- **Developed by**: Ali Sehpar Shikoh - [Eclipse Geomatics Ltd.](https://www.eclipsegeomatics.com)
- **Last Updated**: February 2026
