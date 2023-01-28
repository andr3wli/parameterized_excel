# libraries and variables and etc. ####
library(openxlsx)

# Read in data/little data cleaning ####
paid <- readr::read_csv("data/paid_cash_year1.csv")
received <- readr::read_csv("data/received_cash_year1.csv")

# need to reorder the quarters to match for the left join. Order we want: Q3, Q4, Q1, Q2
q_order <- c("Q3", "Q4", "Q1", "Q2")
location_order <- c("Mountain View Branch", "Fountain Valley Branch", "Pine Rest Branch", "Bayview Community Branch")

received <- received |> 
  dplyr::mutate(location_name = factor(location_name, levels = location_order),
         fiscal_quarter = factor(fiscal_quarter, levels = q_order)) |> 
  dplyr::arrange(fiscal_quarter) |> 
  dplyr::arrange(location_name) |> 
  dplyr::mutate(location_name = as.character(location_name),
         fiscal_quarter = as.character(fiscal_quarter))

# combine the data sets  
data <- cbind(paid, dplyr::select(received, client_payment, other_payment))

dat <- split(data, data$location_name)

# load the excel template ####
wb <- loadWorkbook(here::here("template.xlsx"))

# Create the different worksheets ####
for(i in 1:length(dat)) { # Create the number of sheets equal to the number of unique branch location numbers
  cloneWorksheet(wb, 
                 sheetName = paste0("location_", unique(dat[[i]]$location_number)),
                 clonedSheet = "template")
}

# remove the template 
removeWorksheet(wb, "template")

# Add branch name to the different worksheets ####

# get the unique branch names generated for worksheets to be able to loop through it 
worksheets <- vector("character", length(dat))
for(i in 1:length(dat)) {
  worksheets[[i]] <- paste0("location_", unique(dat[[i]]$location_number))
}

for(i in 1:length(dat)) {
  writeData(wb,
            sheet = worksheets[i],
            x = unique(dat[[i]]$location_name),
            startCol = 2,
            startRow = 2)
}

# Add the quarter dates into worksheets (not pulling from data) ####
title <- c("Jan - Mar 2022", "Apr - Jun 2022", "Jul - Sep 2022", "Oct - Dec 2022", "2021-2022 Total")

for(i in 1:length(dat)) {
  for(j in 1:length(title)) {
    writeData(wb,
              sheet = worksheets[i],
              x = title[j],
              startCol = j + 2,
              startRow = 3)
  }
}

# Measures - items in table that is 4 across - turn this into function when have more time ####
for(i in 1:length(dat)){ 
  item <- dat[[i]]$client_payment # clients and customer
  
  for(k in 1:length(item)) {
    writeData(wb,
              sheet = worksheets[i],
              x = item[k],
              startCol = k + 2,
              startRow = 5)
  }
}

for(i in 1:length(dat)){ 
  item <- dat[[i]]$other_payment # other misc operations 
  
  for(k in 1:length(item)) {
    writeData(wb,
              sheet = worksheets[i],
              x = item[k],
              startCol = k + 2,
              startRow = 6)
  }
}

for(i in 1:length(dat)){ 
  item <- dat[[i]]$inventory # purchase of inventory 
  
  for(k in 1:length(item)) {
    writeData(wb,
              sheet = worksheets[i],
              x = item[k],
              startCol = k + 2,
              startRow = 9)
  }
}

for(i in 1:length(dat)){ 
  item <- dat[[i]]$salary # salary/wages
  
  for(k in 1:length(item)) {
    writeData(wb,
              sheet = worksheets[i],
              x = item[k],
              startCol = k + 2,
              startRow = 10)
  }
}

for(i in 1:length(dat)){ 
  item <- dat[[i]]$income_tax # income tax
  
  for(k in 1:length(item)) {
    writeData(wb,
              sheet = worksheets[i],
              x = item[k],
              startCol = k + 2,
              startRow = 11)
  }
}

for(i in 1:length(dat)){ 
  item <- dat[[i]]$general # general and administrative expenses
  
  for(k in 1:length(item)) {
    writeData(wb,
              sheet = worksheets[i],
              x = item[k],
              startCol = k + 2,
              startRow = 12)
  }
}
# Get the sum for each measure (sum for each row) ####
for(i in 1:length(dat)){ 
  for(row in 1:14) {
    if (row %in% c(1:4, 7, 8, 13)) {
      next
    }
    writeFormula(wb,
                 sheet = worksheets[i],
                 x = paste0("=SUM(C", row, ":", "F", row, ")"),
                 startCol = 7,
                 startRow = row)
  }  
}

# Get the sum for the 4 quarters for cash IN (Column c, D, E, F and Row 5 and 6) ####
for(i in 1:length(dat)){ 
  for(cash_in in seq(from = 3, to = 7, by = 1)) {
    writeFormula(wb,
                 sheet = worksheets[i],
                 x = paste0("=(", LETTERS[cash_in], 5, "+", LETTERS[cash_in], 6, ")"),
                 startCol = cash_in,
                 startRow = 4)
  }
}

# Get the sum for the 4 quarters for the cash OUT (Column C, D, E, F and Row 9:12) #### 
for(i in 1:length(dat)){ 
  for(cash_out in seq(from = 3, to = 7, by = 1)) {
    writeFormula(wb,
                 sheet = worksheets[i],
                 x = paste0("=SUM(", LETTERS[cash_out], 9, ":", LETTERS[cash_out], 12, ")"),
                 startCol = cash_out,
                 startRow = 8)
  }
}

# Get the net income for the business by quarter ####
for(i in 1:length(dat)){ 
  for(net in seq(from = 3, to = 7, by = 1)) {
    writeFormula(wb,
                 sheet = worksheets[i],
                 x = paste0("=(", LETTERS[net], 4, "+", LETTERS[net], 8, ")"),
                 startCol = net,
                 startRow = 14)
  }
}

# Save the worksheet #### 
saveWorkbook(wb, "example_report.xlsx", overwrite = TRUE)
#saveWorkbook(wb, "example_report_2.xlsx", overwrite = TRUE) # use this to save another for examples 

