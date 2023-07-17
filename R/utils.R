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


is_list_empty <- function(list) {
  
  # this function checks wether 
  # the list is empty
  # 
  # returns TRUE if so
  
  all(
    sapply(
      list,
      is.null
    )
  )
  
  
  
  
  
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


# warning messages; ####

.pkg_warning <- function(
    header,
    body
) {
  
  #' function information
  #' 
  #' @param header a character string with the
  #' warning header.
  #' 
  #' @param body a character string with 
  #' warning body.
  
  cli_alert_warning(
    paste(
      
      # 1) warning header
      style_bold(
        col_magenta(
          header
        )
      ),
      
      # 2) warning body
      col_magenta(
        body
      )
    )
  )
  
}


# error messages; ####
.pkg_error <- function(
    header,
    param,
    description
) {
  
  #' function information
  #' 
  #' @param header a character string with
  #' the warning hader
  #' @param param a string with the parameters
  #' causing the error.
  #' @param description a description of the error.
  cli_abort(
    message = paste(
      header,
      param,
      description
    ),
    call = NULL
  )
  
  
  
  
}

.pkg_inform <- function(
    header,
    param,
    description
) {
  
  
  rlang::inform(
    message = paste(
      header,
      param,
      description
    ),
    use_cli_format = TRUE
  )
  
  
}



# end of script; ####