# script: Utilities for the 
# library
# author: Serkan Korkmaz
# date: 2023-05-27
# objective: Generate helper functions 
# to ease the programming
# script start; ####

# group_by helper; ####
# 
# This function groups the data
.group_by <- function(
    DT,
    by
) {
  
  # local variables;
  # 
  # These variables are needed for 
  # further in the function
  
  # 1) get column names
  # for the dcast and melt
  # functions
  DT_colnames <- colnames(
    DT
  )
  
  # 2) extract locations of 
  # character and factor colums
  # to identify id_cols and value.vars
  idx <- which(
    sapply(
      X = DT,
      FUN = function(x) {
        
        # Check if its either character
        # or factor
        (is.character(x) | is.factor(x))
        
      }
    )
  )
  
  # 3) extract relevant measure and
  # id vars
  id.vars <- DT_colnames[
    idx
  ]
  
  # extract all column names
  # not found in id.vars
  measure.vars <- setdiff(
    x = DT_colnames,
    y = id.vars
  )
  
  
  # 1) melt the DT
  # accordingly
  DT <- melt(
    data         = DT,
    value.name   = 'value',
    measure.vars = measure.vars,
    id.vars      = id.vars
  )
  
  # 2) dcast data;
  DT <- dcast(
    data = DT,
    formula = as.formula(
      object = paste(
        collapse = '~',
        c(
          # lhs
          paste(c(by$row,'variable'), collapse = '+'),
          
          # rhs
          paste(by$column, collapse = '+')
        )
        
      )
    ),
    sep = '//',
    value.var = 'value'
  )
  
  
  
  return(
    DT
  )
  
}


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