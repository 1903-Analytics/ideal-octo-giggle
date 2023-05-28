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
#' Generate a Workbook which is also stored in `file`.
#' 
#' @importFrom openxlsx createWorkbook
#' @importFrom openxlsx saveWorkbook
#' 
#' @inheritParams openxlsx::createWorkbook
#' @inheritParams openxlsx::saveWorkbook
#' 
#' @param theme A named list of parameters
#' 
#' @returns A workbook
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
  
  message('Creating workbook:\n')
  
  
  iteration <- 6
  
  pb <- txtProgressBar(
    min = 0,
    max = iteration,
    style = 3,
  )
  
  
  
  # 1) Determine gloabl parameters
  # for the list passed
  # to all functions
  
  # 1.1) Verify structure validity
  # and detmine type;
  # 
  # Can be list (Multiple columns within sheets)
  # or 
  # data.table (A single column within sheets)
  type <- get_type(
    list
  )
  
  
  
  # 1.2) Determine coordinates of each
  # table within the lists
  coordinates <- get_coordinate(
    list = list,
    type = type,
    theme = theme
  )
  
  setTxtProgressBar(
    pb = pb,
    value = 1
  )
  
  # 2) Build the workbook;
  
  # 2.1) Generate workbook
  # TODO: consider wether this
  # has to be moved outside of the 
  # function
  wb <- createWorkbook(
    creator  = creator,
    title    = title,
    subject  = subject,
    category = category
  )
  
  setTxtProgressBar(
    pb = pb,
    value = 2
  )
  
  # 2.2) Generate worksheets
  # based on the list
  add_worksheet(
    wb = wb,
    list = list
  )
  
  setTxtProgressBar(
    pb = pb,
    value = 3
  )
  # 2.2) Add data to the sheets
  # based on the list
  add_data(
    wb = wb,
    type = type,
    coordinate = coordinates,
    list = list,
    theme = theme
  )
  
  setTxtProgressBar(
    pb = pb,
    value = 4
  )
  
  # 2.3) Add a theme to the 
  # data
  add_theme(
    wb = wb,
    coordinates = coordinates,
    type = type,
    theme = list(
      color = theme$color
    )
  )
  
  setTxtProgressBar(
    pb = pb,
    value = 5
  )
  
  # 2.4) Store the workbook
  # locally
  # 
  # NOTE: Overwrite is TRUE by
  # default
  saveWorkbook(
    wb = wb,
    file = file,
    overwrite = overwrite
  )
  
  setTxtProgressBar(
    pb = pb,
    value = 6
  )
  
  # 3) Return statement;
  # 
  # Return the workbook
  # to enable enduser with 
  # possibility of using it further
  return(
    wb
  )
  
}

# end of script; ####