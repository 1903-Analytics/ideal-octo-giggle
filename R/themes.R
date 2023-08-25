# script: themes
# author: Serkan Korkmaz
# date: 2023-05-27
# objective: Generate a function which 
# creates themes for the tables
# script start; ####
#' add_theme
#' 
#' @keywords internal
#' @importFrom openxlsx createStyle
#' @importFrom openxlsx addStyle
#' @importFrom RColorBrewer brewer.pal
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>
themeOps <- function(
    theme
) {
  
  if (missing(theme) | is.null(theme)) {
    
    
    .pkg_warning(
      header = 'No theme detected:',
      body   = 'using color: {.val Greys} compact: {.val FALSE} and combine: {.val FALSE}'
    )
    
    
    theme <- list(
      compact = FALSE,
      combine = FALSE,
      color   = 'Greys'
    )
    
  }
  
  if (is.null(theme$color)) {
    
    theme$color <- 'Greys'
    
  }
  
  if (is.null(theme$compact)) {
    
    theme$compact <- FALSE
    
  }
  
  if (is.null(theme$combine)) {
    
    theme$combine <- FALSE
    
  }
  
  
  
  # Check if color exists
  # otherwise set to default;
  available_colors <- available_colors()
  
  color_indicator <- grepl(
    pattern = theme$color, 
    x = available_colors,
    ignore.case = TRUE
  )
  
  inavailability_indicator <- as.logical(
    sum(color_indicator) == 0 
    )
  
  
  if (inavailability_indicator) {
    
    warning(
      paste0(
        'Could not find color: ', theme$color,'. Setting default.' 
      ),
      call. = FALSE
    )
    
    theme$color <- 'Greys'
    
  } else {
    
    theme$color <- available_colors[color_indicator]
    
  }
  
  # TODO: Needs to be in a 
  # an try function.
  theme$template <- 'default'
  
  
  
  return(
    theme
  )
}






add_theme <- function(
    wb,
    wb_backend,
    theme = list(
      color = 'Reds',
      template = 'default'
    )
) {
  
  
  
  # Extract the theme
  # for the workbook
  foo <- try(
    match.fun(
    FUN = paste0(
      '.theme_', theme$template
  )
  ),
  silent = TRUE
  )
  
  if (inherits(foo, 'try-error')) {
    
    .pkg_warning(
      header = 'Template not found:',
      body   = 'using the {.val default}'
    )
    
    foo <- match.fun(
      '.theme_default'
    )
    
  }
  
  
  
  # NOTE: Usie match.fun or get
  
  # Extract color scheme
  color <- brewer.pal(
    n = 3,
    name = theme$color
  )
  
  
  foo(
    wb = wb,
    wb_backend = wb_backend,
    color = color
  )
  
  
}



#' add_color 
#' 
#' @keywords internal
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