# script: worksheet
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These functions adds worksheets
# to the workbook
# Setup function;
#' data_wb_backend
#' 
#' @importFrom data.table rbindlist
#' @importFrom data.table data.table
#' @importFrom openxlsx mergeCells
#' @importFrom data.table fifelse
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>



.grouping_row <- function(
    wb,
    wb_backend
) {
  
  DT <- rbindlist(
    wb_backend
  )
  
  # 2) add relevant
  # headers
  lapply(
    X   = 1:nrow(DT),
    FUN = function(i) {
      
      # 1) Extract rowwise
      # elements
      DT_ <- DT[i,]
      
      # 1.1) Extract
      # relevant values
      sheet    <- DT_$sheet_id
      startCol <- DT_$x_start
      startRow <- DT_$y_start
      rows     <- startRow:DT_$y_end
      cols     <- startCol:DT_$x_end
      
      
      
      
      addStyle(
        wb = wb,
        sheet = sheet,
        style = createStyle(
          fontColour = 'black',
          valign = 'center',
          halign = 'center',
          wrapText = TRUE
        ),
        rows     = rows,
        cols     = cols,
        gridExpand = TRUE,
        stack = TRUE
      )
      
      # mergeCells(
      #   wb = wb,
      #   sheet = sheet,
      #   rows     = rows,
      #   cols     = startCol
      #   # cols  = DT_$x_start:DT_$x_end,
      #   # rows  = DT_$y_start:DT_$y_end
      # )
      
      
      
    }
  )
  
  
}




.header <- function(
  wb,
  wb_backend,
  fgFill
) {
  
  
  # function information;
  # 
  # adds headers and subheaders
  # with relevant captions
  
  # TODO: how about header DT
  # if you have multiple rows
  # where the headers are not needed
  
  # 1) convert to data.table
  # as the wb_backend is passed as a list
  DT <- rbindlist(
    wb_backend
  )
  
  # 2) add relevant
  # headers
  lapply(
    X   = 1:nrow(DT),
    FUN = function(i) {
      
      # 1) Extract rowwise
      # elements
      DT_ <- DT[i,]
      
      # 1.1) Extract
      # relevant values
      sheet    <- DT_$sheet_id
      startCol <- DT_$x_start
      startRow <- DT_$y_start
      rows     <- startRow:DT_$y_end
      cols     <- startCol:DT_$x_end
      
      # 1) Write caption
      # to the sheet
      writeData(
        wb = wb,
        sheet = sheet, 
        # TODO: Needs a different caption
        # here.
        x = paste0('Title: ', DT$caption),
        startCol = startCol,
        startRow = startRow,
        colNames = FALSE,
        rowNames = FALSE,
        borders = "surrounding",
        borderStyle = "thin"
      )
      
      addStyle(
        wb = wb,
        sheet = sheet,
        style = createStyle(
          fontColour = 'black',
          valign = 'center',
          halign = 'left',
          fontSize = 16,
          wrapText = TRUE,
          textDecoration = 'bold',
          indent = 2,
          fgFill = fgFill[5]
        ),
        rows     = rows,
        cols     = cols,
        # cols  = DT_$x_start:DT_$x_end,
        # rows  = DT_$y_start:DT_$y_start,
        gridExpand = TRUE,
        stack = TRUE
      )
      
      
      
      mergeCells(
        wb = wb,
        sheet = sheet,
        rows     = rows,
        cols     = cols
        # cols  = DT_$x_start:DT_$x_end,
        # rows  = DT_$y_start:DT_$y_end
      )
      
      # 
    }
  )
  
  
}



.grouping_header <- function(
    wb,
    wb_backend,
    fgFill
) {
  
  # function information;
  # 
  # This function generates
  # headers for rows and columns

  # 2) add relevant
  # headers
  lapply(
    X   = 1:nrow(wb_backend),
    FUN = function(i) {
      
      # 1) Extract rowwise
      # elements
      DT_ <- wb_backend$group_coords[[i]]
      sheet_id <- wb_backend$sheet_id[i]
      
      # 2) Split by ownership
      # as to group relevant rows
      DT_list <- split(
        x = DT_,
        f = DT_$ownership
      )
      
      lapply(
        X = DT_list,
        FUN = function(DT) {
          
          # each DT contains seperate information
          # about grouping 
          lapply(
            X = 1:nrow(DT),
            FUN = function(i) {
              
              sheet    <- sheet_id
              startCol <- DT[i,]$x_start
              startRow <- DT[i,]$y_start
              rows     <- startRow:DT[i,]$y_end
              cols     <- startCol:DT[i,]$x_end
              
              
              writeData(
                wb = wb,
                sheet = sheet,
                startCol = startCol,
                startRow = startRow,
                x = DT[i,]$group
              )
              
              
              if (DT[i,]$ownership == 'row') {

                addStyle(
                  wb = wb,
                  sheet = sheet,
                  rows = rows,
                  cols = cols,
                  stack = TRUE,
                  style = createStyle(
                    border = c('right')
                  )
                )

                # Add lines
                test <- DT_[ownership %chin% 'row'] 
                
                # NOTE: If the test data
                # is above two rows then 
                # we need to remove top and bottom
                # to avoid wrong bordering
                indicator <- fifelse(
                  test = nrow(test) > 2,
                  yes  = TRUE,
                  no   = FALSE
                )
                
                if (indicator) {
                  
                  test <- test[-c(.N)]
                  
                } else {
                  
                  test <- test[c(1)]
                  
                  
                }
                
                lapply(
                  1:nrow(test),
                  function(i){
                    addStyle(
                      wb = wb,
                      sheet = sheet,
                      
                      # Based on the rows it might be
                      # extended to fit the entire table.
                      # TODO: Later point
                      rows = test$y_start[i]:test$y_end[i],
                      cols = test$x_start[i]:test$x_end[i],
                      stack = TRUE,
                      style = createStyle(
                        border = c('Bottom'),
                        borderStyle = 'dashed'
                      )
                    )
                  }
                )
                
                

              }
              
              
              # add styles and merge cells
              # 
              # NOTE: These styles adnd merged cells are applied
              # to all the headers; rows and headers
              addStyle(
                wb = wb,
                sheet = sheet,
                style = createStyle(
                  fontColour = 'black',
                  valign = 'center',
                  wrapText = TRUE,
                  fgFill = fgFill,
                  halign = 'center',
                  fontSize = 12,
                  # textDecoration = 'bold',
                  indent = 2
                ),
                rows     = rows,
                cols     = cols,
                gridExpand = TRUE,
                stack = TRUE
              )
              
              mergeCells(
                wb = wb,
                sheet = sheet,
                rows     = rows,
                cols     = cols
              )
              
            }
          )
          
        }
      )
    }
  )
}



# table headers;
table_headers <- function(
    wb,
    wb_backend,
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
  header_dt <- wb_backend[
    between(
      x = column_order,
      lower = min(wb_backend$column_order), 
      upper = max(wb_backend$column_order)
    ) & 
      table_order == 1
  ]
  
  
  header_DT <- rbindlist(
    wb_backend$header_coords
  )
  
  # table headers; 
  lapply(
    1:nrow(header_DT),
    function(i) {
      
      DT <- header_DT[i,]
      
      writeData(
        wb = wb,
        sheet = DT$sheet_id, 
        x = paste0('Title: ', DT$column_caption[1]),
        startCol = DT$x_start,
        startRow = DT$y_start,
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
        cols  = DT$x_start:DT$x_end,
        rows  = DT$y_start:DT$y_start,
        gridExpand = TRUE,
        stack = TRUE
      )



      mergeCells(
        wb = wb,
        sheet = DT$sheet_id,
        cols  = DT$x_start:DT$x_end,
        rows  = DT$y_start:DT$y_end
      )
      
      
    }
  )
  
  
  
  
  
  # table headers; 
  lapply(
    1:nrow(wb_backend),
    function(i) {
      
      DT <- wb_backend[i,]
      
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


.table_caption <- function(wb, wb_backend, fgFill) {
  
  lapply(
    1:nrow(wb_backend),
    function(i) {
      
      # Extract
      DT <- wb_backend$caption_coords[[i]]
      sheet <- wb_backend$sheet_id[i]
      
      # NOTE: There is 
      # a split between labels
      # so it has to be wrapped in lapply
      
      lapply(
        1:nrow(DT),
        function(i) {
          
          addStyle(
            wb = wb,
            sheet = sheet,
            style = createStyle(
              fgFill = fgFill
            ),
            cols = DT[i,]$x_start:DT[i,]$x_end,
            row  = DT[i,]$y_start:DT[i,]$y_end,
            gridExpand = TRUE,
            stack = TRUE
          )
          
          
          writeData(
            wb = wb,
            sheet = sheet,
            x = DT[i,]$label,
            startCol = DT[i,]$x_start,
            startRow = DT[i,]$y_start
          )
          
        }
      )
      
      
    }
  )
  
  
  
}



# add_headers; #####
# 
# add_headers include:
# title headers
# caption headers

add_headers <- function(
    wb,
    wb_backend,
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
  
  # 1) grouping headers;
  .grouping_header(
    wb = wb,
    wb_backend = wb_backend,
    fgFill = fgFill[2]
  )
  
  # NOTE: the caption ONLY
  # colors the caption area.
  .table_caption(
    wb = wb,
    wb_backend = wb_backend,
    fgFill = fgFill[3]
  )
}


