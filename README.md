# Meta-Analysis Automation Script

## Description
This R script performs meta-analyses generating summary statistics, forest plots, leave-one-out sensitivity analyses, Egger’s tests and funnel plots.

## Input
- **Excel file:** `Meta_analysis_data.xlsx`
- The data used for the study, 'A systematic review and meta-analysis of the impact of vaccination on prevention of long COVID', is available in the supplementary section
- Each sheet should contain:
  - `Study`: Study identifier
  - `Ratio`: Odds ratio
  - `LCI`: Lower 95% confidence interval
  - `UCI`: Upper 95% confidence interval
  - `Analysis`: Grouping variable for sub-analyses

## Output
For each sheet, a folder is created containing:
- Meta-analysis summary 
- Forest plot
- Leave-one-out sensitivity plot 
- Egger’s test results *(if ≥10 studies)*
- Funnel plot *(if ≥10 studies)*

## Required Packages & Versions
- `meta` (tested on version 8.2-0)
- `dplyr` (tested on version 1.1.4)
- `openxlsx` (tested on version 4.2.8)
- `readxl` (tested on version 1.4.5)
