# script: worksheet
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These functions adds worksheets
# to the workbook
# Setup function;
#' .add_worksheet
#' 
#' @importFrom openxlsx addWorksheet
#' @inheritParams openxlsx::createWorkbook
#' @author Serkan Korkmaz <serkor1@duck.com>

# script start; ####

add_worksheet <- function(
    wb         = NULL,
    list_names = NULL 
) {
  
  # function globals;
  sheet_name <- list_names
  
  # set names of sheet if
  # none detected
  # if (is.null(sheet_name)) {
  #   
  #   # set warning;
  #   warning(
  #     'No names detected. Setting default.',
  #     call. = FALSE
  #   )
  #   
  #   sheet_name <- paste(
  #     'Sheet', 1:ncol(list)
  #   )
  #   
  # }
  
  invisible(
    lapply(
      sheet_name,
      function(sheetName) {
        
        addWorksheet(
          wb = wb,
          sheetName = sheetName
        )
        
      }
    )
  )
  
  
  
}




initialise_workbook <- function(
    list_names = NULL,
    creator    = NULL,
    title      = NULL,
    subject    = NULL,
    category   = NULL
    ) {
  
  # 1) Create workbook
  wb <- createWorkbook(
    creator  = creator,
    title    = title,
    subject  = subject,
    category = category
  )
  
  # 2) Add worksheet
  add_worksheet(
    wb = wb,
    list_names = list_names
  )
  
  # 3) return
  return(
    wb
  )
  
}


# end of script; ####