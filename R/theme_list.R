# script: theme_list
# date: 2023-06-23
# author: Serkan Korkmaz, serkor1@duck.com
# objective: Generate a range of themes used to write the data
# to excel
# script start;

# theme: default #####

.theme_default <- function(
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
      'black',
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
  
  # Color content; 
  lapply(
    1:nrow(wb_backend),
    function(i) {
      
      # 1) extract the table coordinates as 
      # a data.table
      DT_ <- wb_backend$table_coords[[i]]
      sheet_id <- wb_backend$sheet_id[i]
      
      # 2) start coloring
      
      # 2.1) table content;
      add_color(
        wb           = wb,
        sheet        = sheet_id,
        fgFill       = option_list$fgFill[3],
        # NOTE: we add one because its grouped.
        # needs a more robust approach at some point.
        rows         = (DT_$y_start+1):DT_$y_end,
        cols         = (DT_$x_start + 1):(DT_$x_end - 1),
        border       = option_list$border[3],
        borderColour = option_list$borderColour[3],
        borderStyle  = option_list$borderStyle[3],
        fontColour   = option_list$fontColour[3] 
      )
      
      # 2.2) Column names
      add_color(
        wb           = wb,
        sheet        = sheet_id,
        fgFill       = option_list$fgFill[1],
        # NOTE: we add one because its grouped.
        # needs a more robust approach at some point.
        rows         = (DT_$y_start),
        cols         = (DT_$x_start):(DT_$x_end - 1),
        border       = c('Top', 'Bottom'),
        borderColour = option_list$borderColour[1],
        borderStyle  = option_list$borderStyle[3],
        fontColour   = option_list$fontColour[3] 
      )
      
      
    }
  )
}




# script end;