#' as_workbook_data
#' 
#' 
#' This function prepares the data
#' that is fed into to generate workbook funciton
#' @param DT a data.table
#' @param structure a named list of variable given as characters for the final workbook structure. NULL by default.
#' @param by a named list of grouping variables. Can be NULL. For column grouping >1, the consequtive declared
#' variables will serve as subheaders.
#' 
#' @returns a named list with class 'workbook data' and, if grouped, grouped class.
#' 
#' @example man/examples/exampleWorkbookData.R
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>
#' 
#' @export


# as_workbook_data; ####
as_workbook_data <- function(
    DT,
    structure = list(
      sheet = NULL,
      column = NULL,
      row = NULL
    ),
    by = list(
      column = NULL,
      row = NULL
    )
) {
  
  
  # TODO: Check wether the grouping
  # and structure includes the same variables
  # if so then it might be bugged. Needs a fix.
  
  # 1) Setup globals;
  sheet <- structure$sheet
  column <- structure$column
  row    <- structure$row
 
  # 1) check if data is
  # provided and convert to data.table
  # if so
  if (missing(DT)) {
  
    .pkg_error(
      header = 'Missing argument:',
      param = '{.val DT}',
      description = 'is missing with no default value.' 
    )  
    
    
  } else {
    
    DT <- data.table::copy(DT)
    
  }
  
  
  if (missing(by)) {
    
    by <- list(
      column = NULL,
      row = NULL
    )
    
  }
  
 
  
  
  
  
  
  # Error-handling;
  if (is.null(sheet)) {


    .pkg_warning(
      header = 'No sheet variable detected:',
      body   = 'using single sheet setup.'
    )
    
    
    

    DT[
      ,
      sheet := 'sheet'
      ,
    ]
    
    sheet <- 'sheet'



  }

  if (is.null(column)) {


    .pkg_warning(
      header = 'No colum variable detected:',
      body   = 'using single colum setup.'
    )


    DT[
      ,
      column := 'colum'
      ,
    ]
    
    column <- 'column'

  }


  if (is.null(row)) {

    

    .pkg_warning(
      header = 'No row variable detected:',
      body   = 'using single row setup.'
    )

    DT[
      ,
      row := 'row'
      ,
    ]
    row <- 'row'
    
    
  }


  # Check if passed column names
  # exists
  indicator <-  grepl(
    pattern = paste(
      tolower(colnames(DT)),
      collapse = '|'
    ),
    x = tolower(c(sheet, column, row)),
    fixed = FALSE
  )
  
  


  if (!all(indicator)) {

    cli::cli_abort(
      message = 'The columns: {.val {c(sheet,column,row)[!indicator]}} doesn\'t exist in the data.'
    )

  }
  


  # 1) Split by sheets;
  DT_ <- split(
    DT,
    by = sheet,
    keep.by = FALSE
  )

  # 2) Split by columns
  # within sheets
  DT_ <- lapply(
    DT_,
    function(element) {
      
      
      
      split(
        element,
        by = column,
        keep.by = FALSE
      )

    }
  )

  # 3) Split by rows
  # witin sheets and columns
  DT_ <- lapply(
    X = DT_,
    FUN = function(element) {

      lapply(
        element,
        function(element_) {
          
          
          # TODO: The split
          # function returns a list.
          # We need to add the group_bytest
          # here.
          # BUG

         DT_list <-  split(
            element_,
            by = row,
            keep.by = FALSE
          )
         
         
         # 2) check if the data is
         # grouped
         if (!is_list_empty(by)) {
           
           
           # TODO: It might be a more robust
           # to add this in 
           DT_list <- lapply(
             DT_list,
             function(x) {
               
               .group_by(
                 DT = x,
                 by = by
               )
               
               
               
             }
           )
           
           
           
         }
          
         
         return(
           DT_list
         )
          
          

        }
      )

    }
  )
  
  # set class for later
  # improvement

  class(DT_) <- c(
    class(DT),
    'workbook_data',
    if (!is_list_empty(by)) {'grouped'}
  )
  
  
  return(
    DT_
  )
  
  
  
}