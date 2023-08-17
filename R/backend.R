# script: worksheet
# author: Serkan Korkmaz
# date: 2023-05-25
# objective: These functions adds worksheets
# to the workbook
# TODO: If a named list is not passed
# it should return default values
# TODO: Some of the functionality is repeasted
# it should be one a single function
# which passes into the remainder
# Setup function;
#' data_coordinates
#' 
#' @importFrom data.table rbindlist
#' @importFrom data.table data.table
#' @importFrom openxlsx addWorksheet
#' @importFrom data.table fifelse
#' 
#' @author Serkan Korkmaz <serkor1@duck.com>


get_caption_coords <- function(DT, caption, subcaption, adjust) {
  
  # the DT is the table coords
  DT_ <- as.data.table(DT)
  
  # NOTE: When the rows are grouped
  # it should subtract stuff to avoid 
  # wrong coloring
  subcaption_DT <- data.table(
    x_start = DT_$x_start,
    x_end   = DT_$x_end - adjust,
    # It should start two rows
    # above the table
    # 
    # NOTE: might be it has to start
    # higher to account for column names
    y_start = DT_$y_start - 3,
    y_end   = DT_$y_start - 2,
    label   = subcaption
    
  )
  
  
  caption_DT <- data.table(
    x_start = DT_$x_start,
    x_end   = DT_$x_end- adjust,
    # Move up parallel
    y_start = subcaption_DT$y_start - 1,
    y_end   = subcaption_DT$y_start - 1,
    label   = caption
  )
  
  
  rbind(
    caption_DT,
    subcaption_DT
  )
  
  
}

# extract data region; #####
# 
# 
# This function extracts
# the data cells from the
# data.table

get_table_coords <- function(
    x_ = 3,
    y_ = 10,
    DT,
    grouping_name = 'Placeholder'
) {
  
  #' function information
  #' 
  #' 
  #' @param x_ the starting x coordinate
  #' for the data. Has to be an integer
  #' @param y_ the starting y coordinate
  #' for the data. has to be an integer
  #' @param DT the data.table that needs 
  #' to be written to the workbook
  
  # error handling and checking
  # for missing arguments
  if (missing(DT)) {
    
    cli::cli_abort(
      message = c(
        '{.val DT} must be provided'
      )
    )
    
  }
  
  if (!inherits(DT, 'data.table')) {
    
    cli::cli_abort(
      message = c(
        "{.val DT} must be a data.table",
        "x" = "{.val DT} is {.cls {class(DT)}} class"
      )
    )
    
  }
  
  
  # 1) extract coordinates;
  # TODO: has to be converted to 
  # excel coordinates later
  data.table(
    x_start = x_,
    x_end   = x_ + ncol(DT),
    y_start = y_,
    # NOTE: we might need to subtract one
    # from this.
    y_end   = y_ + nrow(DT)
  )
  
}


# extract grouping variables; #####
# if there is grouping by column and/or
# rows the backend should two entire rows above
# the tables in each

get_grouping_coords <- function(
    base_coordinates,
    grouping_name = 'Placeholder'
) {
  
  
  # error handling and checking
  # for missing arguments
  if (missing(base_coordinates)) {
    
    cli::cli_abort(
      message = c(
        '{.val base_coordinates} must be provided'
      )
    )
    
  }
  
  
  if (!inherits(base_coordinates, 'data.table')) {
    
    cli::cli_abort(
      message = c(
        "{.val base_coordinates} must be a data.table",
        "x" = "{.val base_coordinates} is {.cls {class(base_coordinates)}} class"
      )
    )
    
  }
  
  
  lapply(
    1:nrow(base_coordinates),
    FUN = function(i) {
      
      DT <- base_coordinates$table_content[[i]]
      DT_coords <- base_coordinates$table_coords[[i]]
      
      
      # 1) prepare grouping columns
      # and identify locations
      
      # 1.1) Locate grouping indicators
      # from the column names
      grouping_location <- which(
        sapply(
          X = colnames(DT),
          FUN = grepl,
          pattern = '//'
        )
      )
      
      # 1.2) Extract the names of columns
      # and convert to data.table while
      # adding the locations
      DT_ <- data.table(
        grouped_columns = names(
          grouping_location
        ),
        # NOTE: These have to be altered
        # as they assume a starting x_ = 1
        location =  grouping_location
      )
      
      # 1.3) Split the grouped columns
      # by the seperator
      DT_[
        ,
        # The group is the main group split
        # in small subgroupds
        c('group', 'subgroup') := tstrsplit(
          x = grouped_columns,
          "//",
          fixed = TRUE
        )
        ,
      ]
      
      
      # 
      
      # 1.4) Determine x_start and x_end
      # by using a primitive fifelse approach
      DT_[
        
        ,
        coords := fifelse(
          # If the location is above the previous, then
          # it must be the end of the location
          test = location > shift(location),
          yes  = 'x_end',
          no   = 'x_start',
          na   = 'x_start'
        )
        ,
        by = .(
          group
        )
      ]
      
      # 1.4.1) Test
      DT_ <- DT_[
        ,
        keep := fcase(
          min(location) == location, 1,
          max(location) == location, 1,
          default = 0
        )
        ,
        by = .(
          group
        )
      ][
        keep == 1
      ][
        ,
        keep := NULL
        ,
      ]
      
      # 1.5) dcast the data so locations are read
      # column wise.
      DT_ <- dcast(
        data = DT_,
        formula = group ~ coords,
        value.var = 'location'
      )
      
      # 1.6) Add the remainder columns as
      # category
      # TODO: Has to be named from outside the
      # function
      DT_ <- rbind(
        DT_,
        
        # create a data.table with the
        # remaining rows
        data.table(
          x_start = 1,
          # was x_start
          x_end   = min(DT_$x_start) - 1,
          
          # TODO: find a better place
          group   = grouping_name
        )
      )
      
      
      
      
      
      
      # 1.7) Add indicator of group ownership
      # if column then the the values are 
      # owned by the groups
      # if row then there is group ownership
      DT_[
        ,
        ownership := 'column'
        ,
      ]
      
      DT_[
        ,
        `:=`(
          # NOTE: Its just a parallel
          # of locations
          x_start = x_start + DT_coords$x_start - 1,
          x_end   = x_end  + DT_coords$x_start - 1,
          # NOTE: Set start higher if you want
          # more space
          y_start = DT_coords$y_start - 1,
          y_end   = DT_coords$y_start - 1
        )
        ,
      ]
      
      
      
      # 2) Prepare row grouping
      # if possible
      get_row_group <- colnames(DT)[
        !grepl(
          x = colnames(DT),
          # NOTE: We are looking for all values
          # not called variable or includes a grouping
          # seperator.
          #
          # TODO: Add a relevant prefix here at
          # a later point
          pattern = 'variable|//'
        )
      ]
      
      # 2.1) Extract relevant rows
      # and check if it exists by wrapping it
      # in try
      DT_row <- try(
        DT[
          ,
          .(
            group = get(
              get_row_group
            )
          )
          ,
        ]
      )
      
      # 2.2) If it is not a try-error
      # then the data is grouped by row values
      # as dictated by the 'by' function
      if (!inherits(DT_row, 'try-error')) {
        
        # 2.2.1) Extract the location
        # of the rows
        DT_row[
          ,
          location := which(
            grepl(
              x = group,
              pattern = paste(
                collapse = '|',
                unique(DT_row$group)
              )
            )
          )
          ,
        ]
        
        # 2.2.2) extract the coordinates
        # by using max and min as these
        # spans more rows - makes fifelse
        # unfeasible
        DT_row[
          ,
          coords := fcase(
            location == max(location), 'y_end',
            location == min(location), 'y_start'
          )
          ,
          by = .(
            group
          )
        ]
        
        # 2.2.3) dcast the data
        # as to capture coordiantes
        DT_row <- dcast(
          data = DT_row[!is.na(coords)],
          formula = group ~ coords,
          value.var = 'location'
        )
        
        
        # 2.2.4) Shift the location
        # of the coordiantes as they
        # assume a start of (1,1)
        DT_row[
          ,
          `:=`(
            # Its a single column and therefore
            # starts and ends at the same y coordinate
            x_start = DT_coords$x_start,
            x_end   = DT_coords$x_start,
            y_start = y_start + DT_coords$y_start,
            y_end   = y_end   + DT_coords$y_start
          )
          ,
        ]
        
        
        # 2.2.5) Add ownership
        # to the data
        DT_row[
          ,
          ownership := 'row'
          ,
        ]
        
        
      }
      
      
      
      DT_ <- rbind(
        DT_,
        DT_row,
        fill = TRUE
      )
      
    }
  )
  
  
  
  
  
  
  
  # return(
  #   DT_
  # )
  
}


















#' list coordinates

# list coordinates;
list_backend <- function(
    list,
    theme = list(
      compact = TRUE,
      combine  = FALSE
    )
) {
  
  
  # if grouped
  grouping_distance <- fifelse(
    test = inherits(list, 'grouped'),
    yes  = 1,
    no   = 0
  )
  
  
  # compact;
  distance <- fifelse(
    theme$compact,
    # was yes = 3, no = 4
    yes = 3,
    no = 4
  )
  
  combine <- fifelse(
    theme$combine,
    yes = -distance,
    no = 0
  )
  
  
  # Sheet iterator in
  # the workbook
  sheet_iterator <- 0
  
  
  DT <- rbindlist(
    flatten(
      lapply(
        list,
        function(sheet_element) {
          
          # Iterate sheet;
          sheet_iterator <<- sheet_iterator + 1
          
          # The columns within
          # each workbook; 
          # Has to be reset between
          # each sheet iteration.
          column_iterator <- 0
          
          x_ <- 3
          
          # extract captions
          # of the columns
          column_caption <- names(
            list[[sheet_iterator]]
          )
          
          if (is.null(column_caption)) {
            
            column_caption <- paste(
              'Column', 1:length(list[[sheet_iterator]])
            )
            
          }
          
          
          lapply(
            sheet_element,
            function(column_element) {
              
              # iterate columns
              column_iterator <<- column_iterator + 1
              
              # The tables within each
              # columns;
              # Has to be reset between each column
              table_iterator <- 0
              y_ <- 10
              
              # identify the largest                    # )
              # ),
              # column number wiithin the list
              cols <<- max(
                sapply(column_element, ncol)
              )
              
              
              table_caption <- names(column_element)
              
              if (is.null(table_caption)) {
                
                table_caption <- paste(
                  'Table', 1:length(column_element)
                )
                
              }
              
              # 1) Generate the table coordinates
              # and then the remaining
              DT <- rbindlist(
                lapply(
                  column_element,
                  function(DT) {
                    
                    # iterate;
                    table_iterator <<- table_iterator + 1
                    
                    DT_ <- data.table(
                      sheet_id = sheet_iterator,
                      column_caption = column_caption[column_iterator],
                      table_caption = table_caption[table_iterator],
                      table_content = list(
                        list[[sheet_iterator]][[column_iterator]][[table_iterator]]
                      ),
                      # nrow    = nrow(DT),
                      # ncol    = ncol(DT),
                      table_coords = list(
                        get_table_coords(
                          x_ = x_,
                          y_ = y_ + grouping_distance,
                          DT = DT 
                        )
                      )
                      
                      # table_coords = list(
                      #   data.table(
                      #   x_start = x_,
                      #   x_end   = x_ + ncol(DT),
                      #   y_start = y_ + grouping_distance,
                      #   y_end   = y_ + nrow(DT)
                      # )
                      # ),
                      # header_coords = list(
                      #   data.table(
                      #     caption = column_caption[column_iterator],
                      #     sheet_id = sheet_iterator,
                      #     x_start = x_,
                      #     x_end   = x_ + ncol(DT) - 1,
                      #     y_start = y_ - 4,
                      #     y_end   = y_ - 3
                      #   )
                      # ),
                      # subheader_coords = list(
                      #   data.table(
                      #     caption = table_caption[table_iterator],
                      #     sheet_id = sheet_iterator,
                      #     x_start = x_,
                      #     x_end   = x_ + ncol(DT) - 1,
                      #     y_start = y_ - 2,
                      #     y_end   = y_ - 1 
                      #   )
                      # ),
                      # x_start = x_,
                      # x_end   = x_ + ncol(DT),
                      # y_start = y_ + grouping_distance,
                      # y_end   = y_ + nrow(DT)
                    )
                    
                    # Add captions 
                    # NOTE: this is a derivative of the table
                    # content and needs no external function
                    DT_[
                      ,
                      caption_coords :=  list(
                        get_caption_coords(
                          DT = DT_$table_coords,
                          caption = DT_$column_caption,
                          subcaption = DT_$table_caption,
                          adjust = grouping_distance
                        )
                        
                      )
                      ,
                    ]
                    
                    
                    
                    #column_id <<- column_id + 1
                    y_  <<- y_ + nrow(DT) + distance + combine + grouping_distance
                    
                   
                    
                    
                    
                    
                    
                    return(DT_)
                    
                  }
                )
              )
              
              
              
              
              if (inherits(list, 'grouped')) {
                
                
                DT[
                  ,
                  group_coords := get_grouping_coords(
                    base_coordinates = DT
                  )
                  
                  
                  ,
                ]
                
              }
              
              
              #caption_id <<- caption_id + 1
              
              x_ <<- x_ + cols + 1
              
              return(DT)
              
            }
          )
          
          
          
        }
      )
    )
  )
  
  return(
    DT
  )
  
  
}


wb_backend <- function(
    list,
    theme = list(
      # If compact the distance 
      # between tables are 0,
      # otherwise 1
      compact = TRUE,
      combine = FALSE
    )
) {
  
  # function information
  # 
  # This function returns a list
  # of the same dimensions with 
  # a data.table containing
  # coordinates of the data.tables
  # to be written to the workbook
  
  if (!inherits(list, 'workbook_data')) {
    
    # This error control function
    # is mainly for development and
    # to stop possible errornous further
    # improvements from users.
    
    .pkg_error(
      header = 'Error:',
      param  = '{.val list}',
      description = 'has to be class: {.val workbook_data}'
    )
  }
  
  # generate the back end values
  DT <- list_backend(
    list = list,
    theme = theme
  )
  
  DT[
    ,
    `:=`(
      table_order  = seq_len(.N)
    )
    
    ,
    by = .(
      sheet_id,
      column_caption
    )
  ][
    ,
    `:=`(
      column_order  = fifelse(
        table_order == 1, 1, 0
      )
    )
    ,
  ][
    ,
    `:=`(
      column_order  = cumsum(
        column_order
      )
    )
    ,
    by = .(
      sheet_id
    )
  ]
  
  
  return(
    DT[]
  )
  
}

