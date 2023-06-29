# script: theme_list
# date: 2023-06-23
# author: Serkan Korkmaz, serkor1@duck.com
# objective: Generate a range of themes used to write the data
# to excel
# script start;

# theme: as_is #####

.theme_as_is <- function(
  wb,
  wb_backend,
  color
) {
  
  # This theme prints the data as
  # is directly to the workbook.
  
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
    1:nrow(wb_backend),
    function(i) {
      
      DT_ <- wb_backend[i,]
      
      lapply(
        1:3,
        function(k) {
          
          
          # Extract wb_backend
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




# script end;