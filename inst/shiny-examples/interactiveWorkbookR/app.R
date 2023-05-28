#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)

colors <- data.table::as.data.table(RColorBrewer::brewer.pal.info,keep.rownames = TRUE)

# Define UI for application that draws a histogram
ui <- fluidPage(

    # Application title
    titlePanel("Generate workbooks with R"),
    title = 'Interactive Workbooks', 

    # Sidebar with a slider input for number of bins 
    sidebarLayout(
      
        sidebarPanel(
          
          fileInput(
            "upload",
            "Upload file(s)",
            multiple = TRUE,
            accept = c('.csv', '.xlsx')
          ),
          
            selectInput(
              inputId = 'color_scheme',
              label = 'Choose colorscheme',
              selected = sample(colors[category %chin% 'seq']$rn, size = 1),
              choices = colors[category %chin% 'seq']$rn
            ),
            
            column(
              width = 6,
              checkboxInput(
                inputId = 'compact',
                label = 'Compact',
                value = FALSE
              )
            ),
            
            column(
              width = 6,
              checkboxInput(
                inputId = 'combine',
                label = 'Combine',
                value = FALSE
              )
            ),
          
          br(),
            
          
          
            downloadButton(
              "downloadData",
              "Download workbook",
              style = "width: 100%;"
              )
            
            
          
          
          
          
        ),

        # Show a plot of the generated distribution
        mainPanel(
          mainPanel(
            tableOutput("contents")
          )
        )
    )
)

# Define server logic required to draw a histogram
server <- function(input, output) {

  # Our dataset
  data <- reactive({
    
    
    test <- lapply(
      input$upload$datapath,
      function(x) {
        
        list(
          fread(x)
        )
        
      }
    )
    
    names(test) <- c('fisk','torsk')
    
    return(test)
    
    
  })
  
  output$downloadData <- downloadHandler(

    filename = function() {
      paste("workbook.xlsx", sep="")
    },

    content = function(file) {

        generate_workbook(
        list = data(),
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
    validate(need(ext == "csv" | ext == 'xlsx', "Please upload a csv file"))


    fread(file$datapath)
  })
  
  
}

# Run the application 
shinyApp(ui = ui, server = server)
