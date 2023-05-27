# script: themes
# author: Serkan Korkmaz
# date: 2023-05-27
# objective: Generate a function which 
# creates themes for the tables
# script start; ####
#' themes
#' 
#' 
#' @importFrom openxlsx createStyle
#' @importFrom openxlsx addStyle
#' @importFrom RColorBrewer brewer.pal
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>







# .add_outerborder <- function(
#   coordinates,
#   wb,
#   sheet
# ) {
#   
#   # TODO: Possible bug;
#   # If the tables are unequal length
#   # then the borders wont be aligned
#   
#   # Left surrounding border
#   addStyle(
#     wb = wb,
#     sheet = sheet,
#     stack = TRUE,
#     gridExpand = FALSE,
#     cols = min(coordinates$x_start)-1,
#     rows = min(coordinates$y_start):max(coordinates$y_end),
#     style = createStyle(
#       border = 'Right'
#     )
#   )
#   
#   # Right surrounding Border
#   addStyle(
#     wb = wb,
#     sheet = sheet,
#     stack = TRUE,
#     gridExpand = FALSE,
#     cols = max(coordinates$x_end),
#     rows = min(coordinates$y_start):max(coordinates$y_end),
#     style = createStyle(
#       border = 'Left'
#     )
#   )
#   
#   # Top Surrounding 
#   
#   
# }



add_theme <- function(
    wb,
    coordinates,
    theme = list(
      color = 'Reds'
    )
) {
  
  
  
  # Extract color scheme
  color <- brewer.pal(
    n = 3,
    name = theme$color
  )
  
  # generate options list
  option_list <- list(
    
    # generates coordinate
    # ranges based on locations
    location = c(
      'header',
      'sidebar',
      'table'
    ),
    
    fgFill = c(
      # Header color
      color[3],
      # Siebar color
      color[2],
      # Table color
      color[1]
    ),
    
    fontColour = c(
      'white',
      'black',
      'black'
    ),
    
    border = c(
      'top',
      'right',
      'top'
    ),
    
    borderColour = c(
      'black',
      'black',
      'lightgray'
    ),
    
    borderStyle = c(
      'thin',
      'thin',
      'thin'
    )
    
    
    
  )
  
  sheet_iterator <- 1
  
  # set colors;
  lapply(
    coordinates,
    function(element) {
      
      # Extract data table
      # with coordinates
      DT <- element
      
      # For every row in DR
      # there is headers, sides and
      # tables
      options_iterator <- 1
      
      # Set colors;
      lapply(
        1:nrow(DT),
        function(i) {
          
          # Extract Row
          # from the data
          DT_ <- DT[i,]
          
          lapply(
            1:3,
            function(k){
              
              # Extract coordinates
              coord_range <- .generate_coord_range_color(
                type = option_list$location[k],
                DT = DT_
              )
              
              message(
                option_list$borderColour[k]
              )
              
              add_color(
                wb           = wb,
                sheet        = sheet_iterator,
                fgFill       = option_list$fgFill[k],
                fontColour   = option_list$fontColour[k],
                rows         = coord_range$rows,
                cols         = coord_range$cols,
                border       = option_list$border[k],
                borderColour = option_list$borderColour[k],
                borderStyle  = option_list$borderStyle[k]
              )
              
            }
          )
          
          
          
          
        }
      )
      
      # Set borders;
      
      
      sheet_iterator <<- sheet_iterator + 1
    }
  )
  
  
  
  
}



#' add_color 
#' 
#' 
#' 
#' @importFrom openxlsx addStyle
#' @importFrom openxlsx createStyle

add_color <- function(
    wb,
    sheet,
    fgFill,
    fontColour = 'white',
    rows,
    cols,
    border,
    borderColour, 
    borderStyle  
) {
  
  
  addStyle(
    wb = wb,
    sheet = sheet,
    gridExpand = TRUE,
    cols = cols,
    rows = rows,
    stack = TRUE,
    style = createStyle(
      fontColour   = fontColour,
      fgFill       = fgFill,
      border       = border,
      borderColour = borderColour,
      borderStyle  = borderStyle
    )
  )
  
  
  
}


# end of script; ####