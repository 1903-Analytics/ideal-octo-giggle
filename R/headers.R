# script: worksheet
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These functions adds worksheets
# to the workbook
# Setup function;
#' data_coordinates
#' 
#' @importFrom data.table rbindlist
#' @importFrom data.table data.table
#' @importFrom openxlsx mergeCells
#' @importFrom data.table fifelse
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>

# table headers;
table_headers <- function(
    wb,
    coordinates,
    theme = list(
      color = 'Reds'
    )
    ) {
  
  # parameters;
  color <- brewer.pal(
    n = 6,
    name = theme$color
  )
  
  fgFill <-  c(
    color[5],
    color[4],
    # Header color
    color[3],
    # Siebar color
    color[2],
    # Table color
    color[1]
  )
  
  
  # column headers;
  header_dt <- coordinates[
    between(
      x = column_order,
      lower = min(coordinates$column_order), 
      upper = max(coordinates$column_order)
    ) & 
      table_order == 1
  ]
  
  # table headers; 
  lapply(
    1:nrow(header_dt),
    function(i) {
      
      DT <- header_dt[i,]
      
      writeData(
        wb = wb,
        sheet = DT$sheet_id, 
        x = paste0('Title: ', DT$column_caption[1]),
        startCol = DT$x_start,
        startRow=(DT$y_start-4),
        colNames = FALSE,
        rowNames = FALSE,
        borders = "surrounding",
        borderStyle = "thin"
      )
      
      addStyle(
        wb = wb,
        sheet = DT$sheet_id,
        style = createStyle(
          fontColour = 'black',
          valign = 'center',
          halign = 'left',
          fontSize = 16,
          textDecoration = 'bold',
          indent = 2,
          fgFill = color[5]
        ),
        cols  = DT$x_start:(DT$x_end-1),
        rows  = (DT$y_start-4):(DT$y_start-3),gridExpand = TRUE,stack = TRUE
      )



      mergeCells(
        wb = wb,
        sheet = DT$sheet_id,
        cols  = DT$x_start:(DT$x_end-1),
        rows  = (DT$y_start-4):(DT$y_start-3)
      )
      
      
    }
  )
  
  
  
  
  
  # table headers; 
  lapply(
    1:nrow(coordinates),
    function(i) {
      
      DT <- coordinates[i,]
      
      writeData(
        wb = wb,
        sheet = DT$sheet_id, 
        x = paste0('Table ', DT$table_order,': ',DT$table_caption[1]),
        startCol = DT$x_start,
        startRow=(DT$y_start-2),
        colNames = FALSE,
        rowNames = FALSE,
        borders = "surrounding",
        borderStyle = "thin"
        )
      
      addStyle(
        wb = wb,
        sheet = DT$sheet_id,
        style = createStyle(
          fontColour = 'black',
          valign = 'center',
          halign = 'left',
          indent = 2,
          fgFill = color[4]
        ),
        cols  = DT$x_start:(DT$x_end-1),
        rows  = (DT$y_start-1):(DT$y_start-2),gridExpand = TRUE,stack = TRUE
      )
      
      
      
      mergeCells(
        wb = wb,
        sheet = DT$sheet_id,
        cols  = DT$x_start:(DT$x_end-1),
        rows  = (DT$y_start-1):(DT$y_start-2)
      )
      
      
    }
  )
  
  
}