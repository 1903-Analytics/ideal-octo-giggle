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
# script start; ####
.add_data <- function(
    wb,
    list,
    color
) {
  
  
  # determine type;
  type <- .get_type(
    list = list
  )
  
  
  iteration <- 1
  
  
  if (grepl(pattern = 'list', x = type)) {
    
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
                  startCol = startCol
                  
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
        
        startCol <- 3
        cols <<- max(
          sapply(element, ncol)
        )
        
        lapply(
          element,
          function(DT) {
            
            startRow <- 3
            
            writeData(
              wb = wb,
              x = DT,
              sheet = iteration,
              startRow = startRow,
              startCol = startCol
              
            )
            
            startRow  <<- startRow + nrow(DT) + 1
            
          }
        )
        
        startCol <<- startCol + cols + 1
        
        iteration <<- iteration + 1
      }
    )
    
    
  }
  
  
  
}
# end of script; ####