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
    color,
    theme = list(
      compact = TRUE,
      combine = FALSE
    )
) {
  
  
  # determine type;
  type <- get_type(
    list = list
  )
  
  # get coordinates
  coordinate <- get_coordinate(
    list = list,
    theme = theme
  )
  
  
  
  
  iteration <- 1
  
  
  if (all(grepl(pattern = 'list', x = type))) {
    
    
    message(
      'Is list'
    )
    
    # TODO: migrate as seperate function
    # can be inside
    
    lapply(
      list,
      function(element) {
        
        
        startCol <- 3
        
        lapply(
          element,
          function(element_) {
            
            cols <<- max(
              sapply(element_, ncol)
            )
            
            startRow <- 3
            
            lapply(
              element_,
              function(DT) {
                
                writeData(
                  wb = wb,
                  x = DT,
                  sheet = iteration,
                  startRow = startRow,
                  startCol = startCol,
                  borders = 'surrounding',
                  headerStyle = createStyle(
                    border = c('TopBottom')
                  )
                  
                  
                )
                
                
                startRow  <<- startRow + nrow(DT) + 1
                
              }
            )
            startCol <<- startCol + cols + 1
            
          }
          
          
        )
        
        
        iteration <<- iteration + 1
      }
    )
    
  } else {
    
    
    # TODO: Migrate as
    # seperate function
    
    
    
    lapply(
      list,
      function(element) {
        
        coordinate <- coordinate[[iteration]]
        coordinate_iteration <- 1
        
        startCol <- 3
        startRow <- 3
        # cols <<- max(
        #   sapply(element, ncol)
        # )
        
        lapply(
          element,
          function(DT) {
            
            writeData(
              wb = wb,
              x = DT,
              sheet = iteration,
              startRow = coordinate$y_start[coordinate_iteration],
              startCol = startCol,
              colNames = fifelse(
                test = theme$combine,
                yes = fifelse(
                  coordinate_iteration > 1, no = TRUE, yes = FALSE
                ),
                no = TRUE
              ),
              borders = 'surrounding',
              borderStyle = 'thin',
              headerStyle = createStyle(
                border = c('TopBottom')
              )
            )
            
            coordinate_iteration <<- coordinate_iteration + 1
            #startRow  <<- startRow + nrow(DT) + 1
            
          }
        )
        
        #startCol <<- startCol + cols + 1
        
        iteration <<- iteration + 1
      }
    )
    
    
  }
  
  
  
}
# end of script; ####