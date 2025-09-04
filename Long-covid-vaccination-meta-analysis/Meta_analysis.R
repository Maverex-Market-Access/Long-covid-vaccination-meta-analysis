# Load required libraries
library('meta')
settings.meta(CIbracket = "(", CIseparator = " - ")

library('dplyr')
library(openxlsx)
library(readxl)

# Get sheet names from Excel file
sheet_names <- excel_sheets('Sensitivity analyses A.xlsx')

# Create a folder for each sheet using sanitized sheet name
for (i in 1:length(sheet_names)){
  folder_name <- gsub("≥", "grtreq", gsub("%", "", gsub(">", "grtr", sheet_names[i])))
  dir.create(folder_name)
}

# Loop through each sheet to perform meta-analysis
for (i in 1:length(sheet_names)){
  
  # Read data from current sheet
  dat1 <- readxl::read_xlsx('Sensitivity analyses A.xlsx', sheet = i)
  
  # Assign group IDs based on 'Analysis'
  dat1 <- dat1 %>%
    group_by(Analysis) %>%
    mutate(id = cur_group_id()) %>%
    ungroup()
  
  # Log-transform Odds Ratios and confidence intervals
  dat1$OR <- log(dat1$Ratio)
  dat1$LCI <- log(dat1$LCI)
  dat1$UCI <- log(dat1$UCI)
  
  folder_name <- gsub("≥", "grtreq", gsub("%", "", gsub(">", "grtr", sheet_names[i])))
  
  # Loop through each group (analysis)
  for (j in 1:max(dat1$id)){
    
    # Subset data for current analysis
    dat <- dat1 %>% filter(id == j)
    
    # Calculate standard error from log-transformed CI
    dat$data_se <- ((as.numeric(dat$UCI) - as.numeric(dat$LCI)) / 3.92)
    
    # Perform random-effects meta-analysis
    m.gen <- metagen(
      TE = dat$OR,
      seTE = dat$data_se,
      studlab = dat$Study,
      sm = 'OR',
      fixed = FALSE,
      random = TRUE,
      method.tau = "DL"
    )
    
    # Save summary of meta-analysis to text file
    file_conn <- file.path(folder_name, paste0("Analysis-", paste(gsub("<", "lessthan",
                                                                       gsub("≥", "grtreq",
                                                                            gsub("%", "",
                                                                                 gsub(">", "grtr",
                                                                                      dat$Analysis[1]))))), ".txt"))
    print(writeLines(capture.output(summary(m.gen)), file_conn))
    closeAllConnections()
    
    # Save forest plot of meta-analysis
    png(file.path(folder_name, file = paste0("Analysis-", paste(gsub("<", "lessthan",
                                                                     gsub("≥", "grtreq",
                                                                          gsub("%", "",
                                                                               gsub(">", "grtr", dat$Analysis[1]))))), "-", folder_name, ".png")), 
        width = 2800, height = 2400, res = 300)
    
    forest(m.gen, layout = "JAMA", fontsize = 8, digits = 2,
           colgap.forest.left = "0.5cm", squaresize = 0.4,
           xlab = "Odds Ratio",
           leftlabs = c("Study", "Ratio (95%CI)"), xlab.pos = 0.4)
    dev.off()
    
    # Leave-one-out sensitivity analysis
    x <- metainf(m.gen)
    
    png(file.path(folder_name, file = paste0("LOO-Analysis-", paste(gsub("<", "lessthan",
                                                                         gsub("≥", "grtreq",
                                                                              gsub("%", "",
                                                                                   gsub(">", "grtr", dat$Analysis[1]))))), "-", folder_name, ".png")), 
        width = 2800, height = 2400, res = 300)
    
    forest(x, layout = "JAMA", fontsize = 8, digits = 2,
           colgap.forest.left = "1cm", squaresize = 0.4,
           xlab = "Odds Ratio")
    dev.off()
    
    # Perform Egger's test and save output if enough studies
    if (nrow(dat) >= 10){
      file_conn <- file.path(folder_name, paste0("Analysis-", paste(gsub("<", "lessthan",
                                                                         gsub("≥", "grtreq",
                                                                              gsub("%", "",
                                                                                   gsub(">", "grtr", dat$Analysis[1]))))), "-eggers.txt"))
      print(writeLines(capture.output(metabias(m.gen, method.bias = "linreg", k.min = nrow(dat))), file_conn))
      closeAllConnections()
    }
    
    # Generate funnel plot if enough studies
    if (nrow(dat) >= 10){
      egger_test <- metabias(m.gen, method.bias = "linreg", k.min = nrow(dat))
      p_value <- round(egger_test$pval, 4)
      
      png(file.path(folder_name, file = paste0("Analysis-", paste(gsub("<", "lessthan",
                                                                       gsub("≥", "grtreq",
                                                                            gsub("%", "",
                                                                                 gsub(">", "grtr", dat$Analysis[1]))))), "-", folder_name, ".-funnel.png")), 
          width = 2800, height = 2400, res = 300)
      
      funnel(m.gen, xlab = paste("Odds Ratio (p-value:", p_value, ")"))
      dev.off()
    }
    
  }
}
