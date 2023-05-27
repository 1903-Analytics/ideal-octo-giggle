# script: worksheet
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These functions adds worksheets
# to the workbook
# Setup function;
#' .add_worksheet
#' 
#' @importFrom openxlsx addWorksheet
#' @author Serkan Korkmaz <serkor1@duck.com>

# script start; ####


add_worksheet <- function(
    wb   = NULL,
    list = NULL 
) {
  
  # function globals;
  sheet_name <- names(
    list
  )
  
  # set names of sheet if
  # none detected
  if (is.null(sheet_name)) {
    
    # set warning;
    warning(
      'No names detected. Setting default.',
      call. = FALSE
    )
    
    sheet_name <- paste(
      'Sheet', 1:length(list)
    )
    
  }
  
  invisible(
    lapply(
      1:length(list),
      function(i) {
        
        addWorksheet(
          wb = wb,
          sheetName = sheet_name[i]
        )
        
      }
    )
  )
  
  
  
}


# end of script; ####