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
      
      
      # 1.1) Extract
      # relevant values
      sheet    <- element$sheet_id
      startCol <- element$x_start
      startRow <- element$y_start
      rows     <- startRow:element$y_end
      cols     <- startCol:element$x_end
      
      
      # TODO: 
      # Get DT
      DT <- element$table_content[[1]]
      colnames(DT) <- stringr::str_remove_all(
        string = colnames(DT), pattern = '.+//')
      
      writeData(
        wb = wb,
        x = DT,
        sheet = element$sheet_id,
        colNames = colNames,
        startRow = element$y_start,
        startCol = element$x_start,
        name = paste0(
          'sh', element$sheet_id,
          '_',
          'col', element$column_order,
          '_',
          'tab', element$table_order
          ),
        borders = 'surrounding'
      )
      
      
      
    }
  )
  
  
  
  
}
# end of script; ####