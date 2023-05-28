# script: adddata
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These functions adds data to the
# workbooks.
# 
# function information;
#' .add_data
#' 
#' @importFrom openxlsx writeData
#' @importFrom openxlsx createStyle
#' @importFrom data.table fifelse
# script start; ####
add_data <- function(
    wb,
    list,
    type,
    wb_backend,
    color,
    theme = list(
      compact = TRUE,
      combine = FALSE
    )
) {
  
  # function parameters;
  combine <- theme$combine
  colNames <- !theme$combine
  
  
  
  lapply(
    1:nrow(wb_backend),
    function(i) {
      
      # Extract row of the wb_backend
      # data
      element <- wb_backend[i,]
      
      if (combine) {
        
        colNames <- as.logical(
          element$table_order == 1
        )
        
      }
      
      writeData(
        wb = wb,
        x = element$table_content[[1]],
        sheet = element$sheet_id,
        colNames = colNames,
        startRow = element$y_start,
        startCol = element$x_start,
        borders = 'surrounding'
      )
      
      
    }
  )
  
  
  
  
}
# end of script; ####