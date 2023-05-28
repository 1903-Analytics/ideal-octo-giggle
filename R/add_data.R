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
    coordinate,
    color,
    theme = list(
      compact = TRUE,
      combine = FALSE
    )
) {
  
  
  # # determine type;
  # type <- get_type(
  #   list = list
  # )
  
  # get coordinates
  # TODO: move coordinates outside
  # coordinate <- get_coordinate(
  #   list = list,
  #   theme = theme
  # )
  
  # Iteration is sheet iterator
  iteration <- 1
  combine <- theme$combine
  colNames <- !theme$combine
  
  
  
  lapply(
    1:nrow(coordinate),
    function(i) {
      
      # Extract row;
      element <- coordinate[i,]
      
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
        # startRow = startRow,
        # startCol = startCol,
        startRow = element$y_start,
        startCol = element$x_start,
        borders = 'surrounding',
        headerStyle = createStyle(
          border = c('TopBottom')
        )
      )
      
      
    }
  )
  
  # if (all(grepl(pattern = 'list', x = type))) {
  #   
  #   # TODO: migrate as seperate function
  #   # can be inside
  #   # - Use cooridnates instead of calculating
  #   # everything from bottom
  #   
  #   lapply(
  #     list,
  #     function(element) {
  #       
  #       coordinate_list <- coordinate[[iteration]]
  #       
  #       
  #       startCol <- 3
  #       
  #       # column iterator
  #       col_iterator <- 1
  #       
  #       lapply(
  #         element,
  #         function(element_) {
  #           
  #           cols <<- max(
  #             sapply(element_, ncol)
  #           )
  #           
  #          
  #           
  #           startRow <- 3
  #           coordinate_iteration <- 1
  #           coordinate_ <- coordinate_list[[col_iterator]]
  #           
  #           lapply(
  #             element_,
  #             function(DT) {
  #               
  #               coordinates <- coordinate_[[coordinate_iteration]]
  #               
  #               # Check if colNames
  #               # need to be remove;
  #               if (combine) {
  #                 
  #                 colNames <- !as.logical(
  #                   coordinate_iteration > 1
  #                 )
  #                 
  #               }
  #               
                # writeData(
                #   wb = wb,
                #   x = DT,
                #   sheet = iteration,
                #   colNames = colNames,
                #   # startRow = startRow,
                #   # startCol = startCol,
                #   startRow = coordinates$y_start,
                #   startCol = coordinates$x_start,
                #   borders = 'surrounding',
                #   headerStyle = createStyle(
                #     border = c('TopBottom')
                #   )
  #                 
  #                 
  #               )
  #               
  #               
  #               # startRow  <<- startRow + nrow(DT) + 1
  #               coordinate_iteration <<- coordinate_iteration + 1
  #             }
  #           )
  #           
  #           col_iterator <<- col_iterator + 1
  #           # startCol <<- startCol + cols + 1
  #           
  #         }
  #         
  #         
  #       )
  #       
  #       
  #       iteration <<- iteration + 1
  #     }
  #   )
  #   
  # } else {
  #   
  #   
  #   # TODO: Migrate as
  #   # seperate function
  #   
  #   
  #   
  #   lapply(
  #     list,
  #     function(element) {
  #       
  #       coordinate <- coordinate[[iteration]]
  #       coordinate_iteration <- 1
  #       
  #       startCol <- 3
  #       startRow <- 3
  #       # cols <<- max(
  #       #   sapply(element, ncol)
  #       # )
  #       
  #       lapply(
  #         element,
  #         function(DT) {
  #           
  #           # Check if colNames
  #           # need to be remove;
  #           if (combine) {
  #             
  #             colNames <- !as.logical(
  #               coordinate_iteration > 1
  #             )
  #             
  #           }
  #           
  #           writeData(
  #             wb = wb,
  #             x = DT,
  #             sheet = iteration,
  #             startRow = coordinate$y_start[coordinate_iteration],
  #             startCol = startCol,
  #             colNames = colNames,
  #             # colNames = fifelse(
  #             #   test = theme$combine,
  #             #   yes = fifelse(
  #             #     coordinate_iteration > 1, no = TRUE, yes = FALSE
  #             #   ),
  #             #   no = TRUE
  #             # ),
  #             borders = 'surrounding',
  #             borderStyle = 'thin',
  #             headerStyle = createStyle(
  #               border = c('TopBottom')
  #             )
  #           )
  #           
  #           coordinate_iteration <<- coordinate_iteration + 1
  #           #startRow  <<- startRow + nrow(DT) + 1
  #           
  #         }
  #       )
  #       
  #       #startCol <<- startCol + cols + 1
  #       
  #       iteration <<- iteration + 1
  #     }
  #   )
  #   
  #   
  # }
  
  
  
}
# end of script; ####