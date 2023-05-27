# script: generate_workbook
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: This function gathers all relevant functions
# to create the workbooks.
# 
# function information;
# 
#' generate_workbook
#' 
#' This function generates and collects
#' the workvook
#' 
#' @importFrom openxlsx createWorkbook
#' @importFrom openxlsx saveWorkbook
#' 
#' @returns stores and returns a workbook in the specified
#' file path
#' 
#' @export
# script start; ####

generate_workbook <- function(
    file      = 'workbook.xlsx',
    list      = NULL,
    creator   = NULL,
    title     = NULL,
    subject   = NULL,
    category  = NULL,
    theme     = list(
      compact = TRUE,
      combine = FALSE,
      color   = 'Reds'
    ),
    overwrite = TRUE
) {
  
  
  coordinates <- get_coordinate(
    list = list,
    theme = theme
  )
  
  
  
  # 1) Generate workbook
  # TODO: consider wether this
  # has to be moved outside of the 
  # function
  wb <- createWorkbook(
    creator  = creator,
    title    = title,
    subject  = subject,
    category = category
  )
  
  # 2) Generate worksheets
  # based on the list
  add_worksheet(
    wb = wb,
    list = list
  )
  
  
  # 3) Add data to the sheets
  # based on the list
  add_data(
    wb = wb,
    list = list,
    theme = theme
  )
  
  add_theme(
    wb = wb,
    coordinates = coordinates,
    theme = list(
      color = theme$color
    )
  )
  
  # 4) Store the workbook
  # locally
  saveWorkbook(
    wb = wb,
    file = file,
    overwrite = overwrite
  )
  
  # 5) Return statement;
  # 
  # Return the workbook
  # to enable enduser with 
  # possibility of using it further
  return(
    wb
  )
  
}

# end of script; ####