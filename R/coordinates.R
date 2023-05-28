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
#' @importFrom openxlsx addWorksheet
#' @importFrom data.table fifelse
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>
data_coordinates <- function(
    list,
    theme = list(
      # If compact the distance 
      # between tables are 0,
      # otherwise 1
      compact = TRUE,
      combine = FALSE
    )
) {
  
  # global options;
  
  # compact;
  distance <- fifelse(
    theme$compact,
    yes = 3,
    no = 4
    )
  
  combine <- fifelse(
    theme$combine,
    yes = -distance,
    no = 0
  )
  
  column_caption <- names(list)
  
  
  # Sheet iterator in
  # the workbook
  sheet_iterator <- 0
  
  lapply(
    list,
    function(sheet_element) {
      
      sheet_iterator <<- sheet_iterator + 1
      table_iterator <- 0
      # Start coordinates; 
      
      # The X coordinate 
      # is the start and end of the header
      x_ <- 3
      
      
      # The Y coordinate
      # is the start and end of the table
      y_ <- 7
      
      table_caption <- names(sheet_element)
      
      rbindlist(
        lapply(
          sheet_element,
          function(DT) {
            
            
            
            table_iterator <<- table_iterator + 1
            
            DT_ <- data.table(
              sheet_id = sheet_iterator,
              table_caption = table_caption[table_iterator],
              column_caption = column_caption[sheet_iterator],
              table_content = list(
                list[[sheet_iterator]][[table_iterator]]
              ),
              nrow    = nrow(DT),
              ncol    = ncol(DT),
              x_start = x_,
              x_end   = x_ + ncol(DT),
              y_start = y_,
              y_end   = y_ + nrow(DT)
            )
          
            # NOTE: The '+1' needs to be
            # changed to a dynamic value
            # depending on wether it should be
            # compact or not
            y_ <<- y_ + DT_$nrow + distance + combine
            
            
            return(DT_)
            
          }
        )
      )
      
      
      
    }
  )
  
  
}


#' list coordinates

# list coordinates;
list_coordinates <- function(
    list,
    theme = list(
      compact = TRUE,
      combine  = FALSE
    )
) {
  
  # global options;
  
  
  # compact;
  distance <- fifelse(
    theme$compact,
    yes = 3,
    no = 4
  )
  
  combine <- fifelse(
    theme$combine,
    yes = -distance,
    no = 0
  )
  
  
  # Sheet iterator in
  # the workbook
  sheet_iterator <- 0
  
  lapply(
    list,
    function(sheet_element) {
      
      # Iterate sheet;
      sheet_iterator <<- sheet_iterator + 1
      
      # The columns within
      # each workbook; 
      # Has to be reset between
      # each sheet iteration.
      column_iterator <- 0
      
      x_ <- 3
      
      # extract captions
      # of the columns
      column_caption <- names(
        list[[sheet_iterator]]
      )
      
      
      lapply(
        sheet_element,
        function(column_element) {
          
          # iterate columns
          column_iterator <<- column_iterator + 1
          
          # The tables within each
          # columns;
          # Has to be reset between each column
          table_iterator <- 0
          y_ <- 7
          
          # identify the largest
          # column number wiithin the list
          cols <<- max(
            sapply(column_element, ncol)
          )
          
          
          table_caption <- names(column_element)
          
          DT <- lapply(
            column_element,
            function(DT) {
              
              # iterate;
              table_iterator <<- table_iterator + 1
              
              DT_ <- data.table(
                sheet_id = sheet_iterator,
                column_caption = column_caption[column_iterator],
                table_caption = table_caption[table_iterator],
                table_content = list(
                  list[[sheet_iterator]][[column_iterator]][[table_iterator]]
                ),
                nrow    = nrow(DT),
                ncol    = ncol(DT),
                x_start = x_,
                x_end   = x_ + ncol(DT),
                y_start = y_,
                y_end   = y_ + nrow(DT)
              )
              
              
              #column_id <<- column_id + 1
              y_  <<- y_ + nrow(DT) + distance + combine
              
              
              return(DT_)
              
            }
          )
          
          
          #caption_id <<- caption_id + 1
          
          x_ <<- x_ + cols + 1
          
          return(DT)
          
        }
      )
      
      
      
    }
  )
  
}







get_coordinate <- function(
    list,
    type,
    theme = list(
      # If compact the distance 
      # between tables are 0,
      # otherwise 1
      compact = TRUE,
      combine = FALSE
    )
    
) {
  
  # function information
  # 
  # This function returns a list
  # of the same dimensions with 
  # a data.table containing
  # coordinates of the data.tables
  # to be written to the workbook
  
  # 1) extract type of list
  # TODO: Move out of the function
  # and add as parameter
  # get_type <- get_type(
  #   list
  # )
  
  
  
  if (all(grepl(pattern = 'list', x = type))) {
    
    coordinate_list <- list_coordinates(
      list = list,
      theme = theme
    )
    
  } else {
    
    coordinate_list <- data_coordinates(
      list = list,
      theme = theme
    )
    
    
  }
  
  DT <- rbindlist(
    flatten(
      coordinate_list
    )
  )
  
  # setorder(
  #   DT,
  #   sheet_id,
  #   column_caption,
  #   table_caption
  #   
  # )
  
  DT[
    ,
    `:=`(
      table_order  = seq_len(.N)
    )

    ,
    by = .(
      sheet_id,
      column_caption
    )
  ][
    ,
    `:=`(
      column_order  = fifelse(
        table_order == 1, 1, 0
      )
    )
    ,
  ][
    ,
    `:=`(
      column_order  = cumsum(
        column_order
      )
    )
    ,
    by = .(
      sheet_id
    )
  ]
  
  
  return(
    DT
  )
  
}

