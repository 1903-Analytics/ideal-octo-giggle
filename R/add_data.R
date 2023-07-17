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
      
      # Exttract table
      # coordinates
      table_coords <- element$table_coords[[1]]
      
    
      if (combine) {
        
        colNames <- as.logical(
          element$table_order == 1
        )
        
      }
      
      
      # 1.1) Extract
      # relevant values
      sheet    <- element$sheet_id
      startCol <- table_coords$x_start
      startRow <- table_coords$y_start
      rows     <- startRow:table_coords$y_end
      cols     <- startCol:table_coords$x_end
      
      
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
        startRow = startRow,
        startCol = startCol,
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