# script: themes
# author: Serkan Korkmaz
# date: 2023-05-27
# objective: Generate a function which 
# creates themes for the tables
# script start; ####
#' add_theme
#' 
#' 
#' @importFrom openxlsx createStyle
#' @importFrom openxlsx addStyle
#' @importFrom RColorBrewer brewer.pal
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>

add_theme <- function(
    wb,
    type,
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
  
  lapply(
    1:nrow(coordinates),
    function(i) {
      
      DT_ <- coordinates[i,]
      
      lapply(
        1:3,
        function(k) {
          
          
          # Extract coordinates
          coord_range <- .color_coordinates(
            location = option_list$location[k],
            DT = DT_
          )
          
          
          add_color(
            wb           = wb,
            sheet        = DT_$sheet_id,
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