# script: verifier
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These set of functions determines the 
# the type of list, and wether it is set up validly
# Setup function;
#' .get_type
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>

# script start; ####

.get_type <- function(
    list
) {
  
  
  # 1) Determine validity
  # of the outer list
  indicator <- all(
    sapply(
      list, inherits, 'list'
    )
  ) 
  
  if (!indicator) {
    
    stop(
      'All parent elements has to be lists.',
      call. = FALSE
    )
    
  }
  
  # 1.1) Determine validity
  # of inner lists
  indicator <- all(
    sapply(
      list,
      function(element) {
        inherits(
          element, 'list'
        )
      }
    )
  )
  
  if (!indicator) {
    
    stop(
      'All child elements has to be lists.',
      call. = FALSE
    )
    
  }
  
  
  list_depth <- unique(
    sapply(
      list,
      function(element) {
        sapply(
          element,
          class
        )
      }
    )
  )
  
  if (inherits(list_depth,'list')) {
    
    list_depth <- unique(
      do.call(
        c,
        list_depth
      )
    )
    
  }
  
  return(
    list_depth
  )
}

# end of script; ####