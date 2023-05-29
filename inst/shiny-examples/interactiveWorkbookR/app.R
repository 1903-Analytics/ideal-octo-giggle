# script: interactiveWoorkbookR
# author: Serkan Korkmaz
# date: 2023-05-29
# objective: Create a ShinyApp which automates
# the workbook generation, and displays the power
# of the workbookR
# script start; ####


library(shiny)

# application inputs; ####
# 
# 
# Extract all sequential colors
# from the RColorBrewer library
colors <- data.table::as.data.table(
  RColorBrewer::brewer.pal.info,
  keep.rownames = TRUE
)[category %chin% 'seq']


# ui; #####
ui <- fluidPage(
  
  # Application title
  titlePanel("Generate workbooks with R"),
  title = 'Interactive Workbooks', 
  
  # Sidebar with a slider input for number of bins 
  sidebarLayout(
    
    sidebarPanel = sidebarPanel(
      width = 12,
      
      selectInput(
        inputId = 'color_scheme',
        label = 'Choose colorscheme',
        selected = sample(
          colors$rn,
          size = 1
          ),
        choices = colors$rn
      ),
      
      # display chosen color schemes;
      plotOutput(
        "plotColorScheme",
        fill = TRUE,
        width = '100%'
        ),
      
      fileInput(
        "upload",
        "Upload file(s)",
        multiple = TRUE,
        accept = c('.csv', '.xlsx')
      ),
      
      div(
        style = 'display: flex; text-align: center;',
        checkboxInput(
          inputId = 'compact',
          label = 'Compact',
          value = FALSE,width = '100%'
        ),
        checkboxInput(
          inputId = 'combine',
          label = 'Combine',
          value = FALSE,width = '100%'
        )
      ),
      
      downloadButton(
        "downloadData",
        "Download workbook",
        style = "width: 100%;"
      )
      
    ),
    
    # Keep the mainPanel empty
    # as there is no need for any displays
    # in this application;
    mainPanel = mainPanel(
      mainPanel(
        
      )
    )
  )
)

# server; ####
server <- function(input, output) {
  
  # Our dataset
  DT <- reactive({
    
    
     rekt <- lapply(
      input$upload$datapath,
      function(file) {
        
        list(
          data.table::fread(
            input = file
            )
        )
        
      }
    )
     
    
     return(
       rekt
     )
    
    
  })
  
  output$downloadData <- downloadHandler(
    
    filename = function() {
      paste("workbook.xlsx", sep="")
    },
    
    content = function(file) {
      
      generate_workbook(
        list = DT(),
        file = file,
        overwrite = TRUE,
        theme = list(
          color = input$color_scheme,
          compact = input$compact,
          combine = input$combine
        )
      )
    }
    
  )
  
  output$contents <- renderTable({
    file <- input$upload
    ext <- tools::file_ext(file$datapath)
    
    req(file)
    validate(need(ext == "csv" | ext == 'xlsx', "Please upload a csv or xlsx file"))
    
    
    fread(file$datapath)
  })
  
  output$plotColorScheme <- renderPlot(
    {
      # Extract  chosen colors;
      chosen_colors <- RColorBrewer::brewer.pal(
        n = 3, 
        name = input$color_scheme
      )
      
      # Create graph
      
      par(bg = '#f5f5f5')
      barplot(
        height = c(1,1,1),
        space = c(0,0,0),
        horiz = TRUE,
        col = c(
          chosen_colors[1],
          chosen_colors[2],
          chosen_colors[3]
          
        ),
        yaxt = 'n',
        xaxt = 'n'
      )
      
      text(
        x = 0.5,
        y = 0.5,
        'body'
      )
      
      text(
        x = 0.5,
        y = 1.5,
        'sidebar'
      )
      
      
      text(
        x = 0.5,
        y = 2.5,
        'caption'
      )
      
    }
  )
}

# Run the application;
shinyApp(
  ui = ui,
  server = server
  )


# end of script; ####

