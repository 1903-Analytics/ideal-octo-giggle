# script: Example of as_workbook_data usage
# author: Serkan Korkmaz
# date: 2023-08-25
# objective: Demonstrate how to use as_workbook_data
# using mtcars
# script start; ####

# prelims;

# garbage collection;
invisible(
  {
    rm(list=ls()); gc()
  }
)

# load libraries
library(data.table)
library(workbookR)


# 1) prepare the data
# in a way that adheres to grouped
# and non-grouped structures
mtcars <- as.data.table(
  mtcars
)

# 1.1) convert variables
# to readable characters
mtcars[
  ,
  `:=`(
    cyl  = paste(cyl, 'cylinder(s)'),
    gear = paste(gear, 'gear(s)'),
    am   = fcase(
      default = 'Automatic',
      am == 1, 'Manual'
    )
  )
  ,
]

# 1.2) aggregate the data
# accross variables to
mtcars <- mtcars[
  ,
  lapply(
    .SD,
    mean
  )
  ,
  by = .(
    cyl,
    gear,
    am
  )
]


# 2) convert to workbook data
# using as_workbook_data

# 2.1) Grouped workbook data
wb_data <- as_workbook_data(
  DT = mtcars,
  by = list(
    column = c('gear', 'am'),
    row    = c('cyl')
  )
)

# class
class(wb_data)

# 2.2) 'as is' with manual and automatic
# gear data in seprate sheets
wb_data <- as_workbook_data(
  DT = mtcars,
  structure = list(
    sheet = 'am'
  )
)

# class
class(wb_data)



# NOTE: the generate_workbook will handle the rest.
# However, if you need to change column names and such,
# you should do it on workbook data

# end of script; ####