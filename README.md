
# Parameterized Excel reports 

The goal of this repo is simple: to generate reproducible reports with Excel to be sent to clients and stakeholders.   

This repo was inspired by [Parameterized reporting](https://bookdown.org/yihui/rmarkdown/parameterized-reports.html) in RMarkdown. Many companies/clients require an Excel spreadsheet, 

**NOTE:** The data was randomly generated with the `runif` function. The template was inspired by the cash flow statement excel template from [ExcelDataPro](https://exceldatapro.com/cash-flow-statement/). The branch names were randomly generated (slightly modified) by [Fantasy Name Generators](https://www.fantasynamegenerators.com/). This is just an example of how to generate excel spreadsheet reports these numbers are not reflective of anything else! 

As well, this example requires the data to be tidy. In a similar task, I queried and cleaned the data before putting it in this script. 

## Requirements

 The code was written with R version 4.2.1. The Excel template and output was created on Version 2212. Lastly, the package `openxlsx` was used to write and edit the Excel worksheets. You can install via cran:
 
 ```
 install.packages("openxlsx")
 ```

## How to generate multiple worksheets 

* Read in and join all the data sources. 