# script: Utilities for the 
# library
# author: Serkan Korkmaz
# date: 2023-05-27
# objective: Generate helper functions 
# to ease the programming
# script start; ####


available_colors <- function() {
  
  
  data.table::as.data.table(
    RColorBrewer::brewer.pal.info,
    keep.rownames = TRUE
  )[category %chin% 'seq']$rn
  
}


flatten <- function(list) {
  
  if (!inherits(list, "list")) {
    
    list <- list(list)
    
  } else {
    
    list <- unlist(
      c(
        lapply(list, flatten)
      ),
      recursive = FALSE
    )
  }
  
  return(
    list
  )
  
}


.color_coordinates <- function(
    location,
    DT
) {
  
  # this function extracts the coloring
  # coordinates of the tables
  
  if (grepl(pattern = 'header', x = location)) {
    
    rows <- DT$y_start
    cols <- DT$x_start:(DT$x_end-1)
    
  }
  
  if (grepl(pattern = 'sidebar', x = location)) {
    
    rows <- (DT$y_start+1):DT$y_end
    cols <- DT$x_start
    
    
  }
  
  if (grepl(pattern = 'table', x = location)) {
    
    cols = (DT$x_start+1):(DT$x_end-1)
    rows = (DT$y_start+1):(DT$y_end)
    
  }
  
  return(
    list(
      rows = rows,
      cols = cols
    )
  )
  
}


# end of script; ####