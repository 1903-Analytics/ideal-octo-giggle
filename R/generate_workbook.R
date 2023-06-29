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
#' 
#' @param theme A named list of parameters
#' @param list A list of data.tables, can be nested. Or created using the as_workbook_data function
#' 
#' @example man/examples/basic_usage.R
#' 
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
    theme     = NULL,
    overwrite = TRUE
) {
  
  # 1) Determine gloabl parameters
  # for the list passed
  # to all functions
  
  # 1.1) Verify structure validity
  # and detmine type;
  # 
  # Can be list (Multiple columns within sheets)
  # or 
  # data.table (A single column within sheets)
  # or
  # grouped
  # 
  # 
  # TODO: Deprecated
  # type <- get_type(
  #   list
  # )
  
  # 1.2) Extract theme-defaults if not
  # supplied
  theme <- themeOps(
    theme = theme
  )
  
  # 1.3) Determine coordinates of each
  # table within the lists
  wb_backend <- wb_backend(
    list = list,
    theme = theme
  )
  
  # 2) Build the workbook;
  
  wb <- initialise_workbook(
    list_names  = names(list),
    creator     = creator,
    title       = title,
    subject     = subject,
    category    = category
  )
  
  # 2.2) Add data to the sheets
  # based on the list
  add_data(
    wb = wb,
    wb_backend = wb_backend,
    list = list,
    theme = theme
  )
  
  
  # 2.3) Add a theme to the 
  # data
  add_theme(
    wb = wb,
    wb_backend = wb_backend,
    theme = theme
  )
  
  # table_headers(
  #   wb = wb,
  #   wb_backend = wb_backend,
  #   theme = theme
  # )
  
  add_headers(
    wb = wb,
    wb_backend = wb_backend,
    theme = theme
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