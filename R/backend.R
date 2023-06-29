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
#
data_backend <- function(
    list,
    theme = list(
      # If compact the distance 
      # between tables are 0,
      # otherwise 1
      compact = TRUE,
      combine = FALSE
    )
) {
  
  
  
  
  # global options;
  
  
  
  
  # compact;
  distance <- fifelse(
    theme$compact,
    yes = 3,
    no = 4
    )
  
  combine <- fifelse(
    theme$combine,
    yes = -distance,
    no = 0
  )
  
  column_caption <- names(list)
  
  
  if (is.null(column_caption)) {
    
    column_caption <- paste(
      'Column', 1:length(list)
    )
    
  }
  
  
  
  
  # Sheet iterator in
  # the workbook
  sheet_iterator <- 0
  
  lapply(
    list,
    function(sheet_element) {
      
      sheet_iterator <<- sheet_iterator + 1
      table_iterator <- 0
      # Start coordinates; 
      
      # The X coordinate 
      # is the start and end of the header
      x_ <- 3
      
      
      # The Y coordinate
      # is the start and end of the table
      y_ <- 7
      
      table_caption <- names(sheet_element)
      
      if (is.null(table_caption)) {
        
        table_caption <- paste(
          'Table', 1:length(sheet_element)
        )
        
      }
      
      rbindlist(
        lapply(
          sheet_element,
          function(DT) {
            
            
            
            table_iterator <<- table_iterator + 1
            
            DT_ <- data.table(
              sheet_id = sheet_iterator,
              table_caption = table_caption[table_iterator],
              column_caption = column_caption[sheet_iterator],
              table_content = list(
                list[[sheet_iterator]][[table_iterator]]
              ),
              nrow    = nrow(DT),
              ncol    = ncol(DT),
              x_start = x_,
              x_end   = x_ + ncol(DT),
              y_start = y_,
              y_end   = y_ + nrow(DT)
            )
          
            # NOTE: The '+1' needs to be
            # changed to a dynamic value
            # depending on wether it should be
            # compact or not
            y_ <<- y_ + DT_$nrow + distance + combine
            
            
            return(DT_)
            
          }
        )
      )
      
      
      
    }
  )
  
  
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
              
              # identify the largest
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
              
              DT <- lapply(
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
                    nrow    = nrow(DT),
                    ncol    = ncol(DT),
                    table_coords = list(
                      data.table(
                      x_start = x_,
                      x_end   = x_ + ncol(DT),
                      y_start = y_ + grouping_distance,
                      y_end   = y_ + nrow(DT)
                    )
                    ),
                    header_coords = list(
                      data.table(
                        caption = column_caption[column_iterator],
                        sheet_id = sheet_iterator,
                        x_start = x_,
                        x_end   = x_ + ncol(DT) - 1,
                        y_start = y_ - 4,
                        y_end   = y_ - 3
                      )
                    ),
                    subheader_coords = list(
                      data.table(
                        caption = table_caption[table_iterator],
                        sheet_id = sheet_iterator,
                        x_start = x_,
                        x_end   = x_ + ncol(DT) - 1,
                        y_start = y_ - 2,
                        y_end   = y_ - 1 
                      )
                    ),
                    x_start = x_,
                    x_end   = x_ + ncol(DT),
                    y_start = y_ + grouping_distance,
                    y_end   = y_ + nrow(DT)
                  )
                  
                  
                  
                  
                  
                  if (inherits(list, 'grouped')) {
                    
                   
                    
                    # Grouping the rows;
                    # TODO: Add row 
                    get_col <- colnames(
                      DT
                    )
                    
                    
                    get_col <- get_col[
                      !grepl(
                        pattern = 'variable|//',
                        x = get_col 
                      )
                    ]
                    
                    
                    grouping_rows <- try(DT[
                      ,
                      .(
                        row_group = get(get_col)
                      )
                      ,
                    ])
                    
                    if (!inherits(idx, 'try-error')) {
                      
                      
                      
                      grouping_rows <- grouping_rows[
                        ,
                        location := which(
                          grepl(
                            x = row_group,
                            pattern = paste(
                              collapse = '|',
                              unique(grouping_rows$row_group)
                            )
                          )
                        )
                        ,
                      ][
                        ,
                        coords := fcase(
                          location == max(location), 'y_end',
                          location == min(location), 'y_start'
                        )
                        ,
                        by = .(
                          row_group
                        )
                      ][
                        !is.na(coords)
                      ]
                      
                      
                      
                      grouping_rows <- dcast(
                        data = grouping_rows,
                        formula = row_group ~ coords,
                        value.var = 'location'
                      )
                      
                      
                      grouping_rows[
                        ,
                        `:=`(
                          sheet_id = sheet_iterator,
                          x_start = x_,
                          x_end   = x_,
                          y_start = y_start + y_ + 1,
                          y_end   = y_end + y_ + 1
                        )
                        ,
                      ]
                      
                      
                      DT_[
                        ,
                        `:=`(
                          grouprow_coords = list(
                            grouping_rows
                          )
                        )
                        ,
                      ]
                      
                      
                    }
                    
                    
                    
                    
                    
                    
                    
                    
                    grouping_indicator <- which(
                      sapply(
                        colnames(DT),
                        grepl,
                        pattern = '//'
                      )
                    )
                    
                  
                    
                    grouping_coords <- data.table(
                      value_cols = names(grouping_indicator),
                      location   = grouping_indicator + x_ - 1,
                      sheet_id   = sheet_iterator
                    )
                    
                    
                    grouping_coords[
                      ,
                      c('group', 'subgroup') := tstrsplit(
                        value_cols,
                        '//',
                        fixed = TRUE
                      )
                      ,
                    ][
                      ,
                      `:=`(
                        coords = fifelse(
                          location > shift(location),
                          yes = 'x_end',
                          no  = 'x_start',
                          na  = 'x_start' 
                        ) 
                      )
                      ,
                      by = .(
                        group
                      )
                    ]
                    
                    grouping_coords <- dcast(
                      grouping_coords,
                      formula = sheet_id + group ~ coords,
                      value.var = 'location'
                    )
                    
                    
                    grouping_coords[
                      ,
                      `:=`(
                        y_start = y_,
                        y_end   = y_ - grouping_distance
                      )
                      ,
                    ]
                    
                    DT_[
                      ,
                      `:=`(
                        group_coords = list(
                          grouping_coords
                        )
                      )
                      ,
                    ]
                    
                  }
                  
                  #column_id <<- column_id + 1
                  y_  <<- y_ + nrow(DT) + distance + combine + grouping_distance
                  
                  
                  
                  
                  
                  
                  return(DT_)
                  
                }
              )
              
              
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
  
  
  DT <- list_backend(
    list = list,
    theme = theme
  )
  
  
  # DT <- rbindlist(
  #   flatten(
  #     coordinate_list
  #   )
  # )
  
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
  # 
  # # add headers;
  # # TODO: Might be a good idea 
  # # to move outside
  # is_grouped <- inherits(
  #   list,
  #   what = 'grouped'
  # )
  # 
  # 
  # if (is_grouped) {
  #   
  #   
  #   
  # }
  
  
  return(
    DT[]
  )
  
}

