# script start; ####

# gc
invisible(
  {
    rm(list = ls()); gc()
  }
)

# load libraries
library(data.table)
library(workbookR)

# 1) prepare the data;
mtcars <- as.data.table(
  mtcars
)

# 1.1) convert relevant
# variables that determines
# the split of the data
mtcars[
  ,
  `:=`(
    cyl_   = paste(cyl, 'cylinder(s)'),
    gear_  = paste(gear, 'Gears'),
    am_    = fcase(
      default = 'Automatic',
      am == 1, 'Manual'
    )
  )
  ,
]


# 2) generate workbook_data
# using as_workbook_data
workbook_data <- as_workbook_data(
  DT    = mtcars,
  # splitting by sheets
  sheet = 'am_',
  # splitting each data within 
  # the sheets by number of
  # gears
  colum = 'gear_',
  # Splitting all data
  # between number of cylinders
  row   = 'cyl_'
)

# 3) generate the workbook
# and store in file
wb <- generate_workbook(
  file = 'workbook.xlsx',
  list = workbook_data
)


# end of script; ####