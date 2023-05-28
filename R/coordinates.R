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
    yes = 1,
    no = 2
    )
  
  combine <- fifelse(
    theme$combine,
    yes = -distance,
    no = 0
  )
  
  
  list_id <- 1
  
  lapply(
    list,
    function(element) {
      
      iteration <- 1
      
      # Start coordinates; 
      
      # The X coordinate 
      # is the start and end of the header
      x_ <- 3
      
      # The Y coordinate
      # is the start and end of the table
      y_ <- 3
      
      DT <- rbindlist(
        lapply(
          element,
          function(element_) {
            
            DT <- data.table(
              list_id = list_id,
              table   = iteration,
              nrow    = nrow(element_),
              ncol    = ncol(element_),
              x_start = x_,
              x_end   = x_ + ncol(element_),
              y_start = y_,
              y_end   = y_ + nrow(element_)
            )
            
            
            
            
            # NOTE: The '+1' needs to be
            # changed to a dynamic value
            # depending on wether it should be
            # compact or not
            y_ <<- y_ + DT$nrow + distance + combine
            iteration <<- iteration + 1
            
            return(DT)
            
          }
        )
      )
      
      list_id <<- list_id + 1
      
      
      return(
        DT
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
    yes = 1,
    no = 2
  )
  
  combine <- fifelse(
    theme$combine,
    yes = -distance,
    no = 0
  )
  
  # list id;
  # 
  # sheet numbers
  
  iteration <- 1
  list_id <- 0
  
  lapply(
    list,
    function(element) {
      list_id <<- list_id + 1
      x_ <- 3
      
      DT <- lapply(
        element,
        function(element_) {
          
          table_id <- 1
          cols <<- max(
            sapply(element_, ncol)
          )
          
          y_ <- 3
          
          
          DT <- lapply(
            element_,
            function(DT) {
              
              
              DT_ <- data.table(
                list_id = list_id,
                table   = table_id,
                nrow    = nrow(DT),
                ncol    = ncol(DT),
                x_start = x_,
                x_end   = x_ + ncol(DT),
                y_start = y_,
                y_end   = y_ + nrow(DT)
              )
              
              
              
              y_  <<- y_ + nrow(DT) + 1
              table_id <- table_id + 1
              return(DT_)
              
            }
          )
          
          
          
          x_ <<- x_ + cols + 1
          
          return(DT)
          
        }
      )
      
      
      iteration <<- iteration + 1
      
      
      return(DT)
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
    
    coordinates <- list_coordinates(
      list = list,
      theme = theme
    )
    
  } else {
    
    coordinates <- data_coordinates(
      list = list,
      theme = theme
    )
    
    
  }
  
  return(
    coordinates
  )
  
}

