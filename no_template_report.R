# libraries and variables and etc. ####
library(openxlsx)

# Read in data/little data cleaning ####
final_df <- readr::read_csv("data/example_df.csv")

# Create the excel sheet ####
wb <- createWorkbook()

# Add different worksheets and create the titles, column names ####

# add worksheets based on what location are in that quarter 
location_num <- unique(final_df$location_number)

for (i in location_num) {
  addWorksheet(wb, 
               sheetName = i)
}

# need to chagne location_num to be sting in order to work...
location_str <- as.character(location_num)

# Create the styles for the title 
header_style <- createStyle(textDecoration = "bold", fgFill = "#c6dbef")
col_name_stlye <- createStyle(textDecoration = "bold", border = "bottom", borderStyle = "thin")

# Create the title for worksheets 
for (i in seq_along(location_str)) {
  location_name <- switch(location_str[i], 
                        "85555" = "Mountain View Branch (85555)",
                        "45836" = "Fountain Valley Branch (45836)",
                        "86788" = "Pine Rest Branch (86788)",
                        "74987" = "Bayview Community Branch (74987)")
  
  # Add the style to the title section 
  addStyle(wb, sheet = location_str[i], style = header_style, cols = 1:5, rows = 2)
  
  # Write the location name to cell 2,2 of the worksheet
  writeData(wb, 
            sheet = location_str[i], 
            x = location_name, 
            startCol = 1, 
            startRow = 2)
}

# Add Column names for the worksheet  
excel_col <- c("Expense", "Previous Quarter", "Current Quarter", "Quarter Difference", "% Difference")

for (i in seq_along(location_str)) {
  for (k in seq_along(excel_col)) {
    
    addStyle(wb, sheet = location_str[i], style = col_name_stlye, cols = 1:5, rows = 3)
    
    writeData(wb, 
              sheet = location_str[i], 
              x = excel_col[k], 
              startCol = k,
              startRow = 3) 
  }
}

# Add columns we have in final_df: model, current_q, and previous_q ####

# add for expense col
for (i in seq_along(location_str)) {
  writeData(wb,
            sheet = location_str[i],
            x = unique(final_df$Model),
            startCol = 1,
            startRow = 4)
}

# manually add the net model 
for (i in seq_along(location_str)) {
  writeData(wb,
            sheet = location_str[i],
            x = "Net Model Income",
            startCol = 1,
            startRow = 11)
}

# add previous quarter
for (i in seq_along(location_str)) {
  writeData(wb,
            sheet = location_str[i],
            x = subset(final_df, location_number == location_str[i])[[4]],
            startCol = 2,
            startRow = 4)
}

# add current quarter
for (i in seq_along(location_str)) {
  writeData(wb,
            sheet = location_str[i],
            x = subset(final_df, location_number == location_str[i])[[5]],
            startCol = 3,
            startRow = 4)
}

# Create the new formula columns based on the prev and current quarter ####

# create the difference column - formula example: =C4-B4 
for (i in seq_along(location_str)) {
  for (k in seq(from = 4, to = 10, by = 1)) {
    writeFormula(wb,
                 sheet = location_str[i],
                 x = paste0("=C", k, "-B", k),
                 startCol = 4,
                 startRow = k)
  }
}

# Create the net model income row 
for (i in seq_along(location_str)) {
  for (k in seq(from = 2, to = 4, by = 1)) {
    writeFormula(wb,
                 sheet = location_str[i],
                 x = paste0("=SUM(", LETTERS[k], "5:", LETTERS[k], "9) - ", LETTERS[k], "10"),
                 startCol = k,
                 startRow = 11)
  }
}


# create the % difference column - formula example: =C4/B4-1
for (i in seq_along(location_str)) {
  for (k in seq(from = 4, to = 11, by = 1)) {
    writeFormula(wb,
                 sheet = location_str[i],
                 x = paste0("=C", k, "/B", k, "-1"),
                 startCol = 5,
                 startRow = k)
  }
}

# Additional style changes before client ready ####

# make the columns wider
for (i in seq_along(location_str)) {
  for (width in seq_along(1:5)) {
    setColWidths(wb,
                 sheet = location_str[i],
                 cols = width,
                 widths = "auto")
  }
}

comma_format <- createStyle(numFmt = "NUMBER")
dollar_format <- createStyle(numFmt = "CURRENCY")

for (i in seq_along(location_str)) {
  for (p in seq(from = 2, to = 4, by = 1)) {
    # n patient column
    addStyle(wb,
             sheet = location_str[i],
             style = comma_format,
             cols = p,
             rows = 4)
    # for the dollar cols
    addStyle(wb,
             sheet = location_str[i],
             style = dollar_format,
             cols = p,
             rows = 5:11)
  }
}

# change the percentage difference column to be percentage number format
percent_format <- createStyle(numFmt = "PERCENTAGE")

for (i in seq_along(location_str)) {
  addStyle(wb,
           sheet = location_str[i],
           style = percent_format,
           cols = 5,
           rows = 4:11)
}

# Add color and border for net income
# also make the single cell for net income bold
income_style <- createStyle(border = "top", borderStyle = "thin", fgFill = "#c6dbef")
income_cell_style <- createStyle(textDecoration = "bold")

for (i in seq_along(location_str)) {
  addStyle(wb,
           sheet = location_str[i],
           style = income_style,
           cols = 1:5,
           rows = 11,
           stack = TRUE)
  
  addStyle(wb,
           sheet = location_str[i],
           style = income_cell_style,
           cols = 1,
           rows = 11,
           stack = TRUE)
}

# save the workbook #### 
saveWorkbook(wb, 
             here::here("output/no_template_report.xlsx"),
             overwrite = TRUE)
