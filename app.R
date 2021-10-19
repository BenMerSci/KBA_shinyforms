library(shiny)
library(bslib)
library(tidyverse)
library(magrittr)
library(WordR)
library(flextable)
library(officer)
library(openxlsx)
library(lubridate)
library(janitor)
library(googledrive)
library(emo)

# Authenticate googledrive
# Informations to authenticate are stored in an .Rprofile
googledrive::drive_auth()

# User interface
ui <- fluidPage(

tags$head(tags$style(
    HTML('
         #sidebar {
            background-color: F9F2E5;
        }

        body, label, input, button, select { 
          font-family: "vivaldi";
        }')
  )),

titlePanel(
fluidRow(
    column(11, h1("Creation of KBA summaries")),
    column(1, img(src = "Canada_KBA_transparent.png", height = 80, width = 120, href = "https://www.kbacanadawiki.org"))
    
  )
),
  theme = bs_theme("progress-bar-bg" = "orange", version = 5),

  sidebarLayout(

    sidebarPanel(id = "sidebar", width = 4,  
      h3("ReadMe", align = "center"),
      "This Shiny application serves to convert Excel KBA proposals into Word KBA summaries for expert revisions.",
      hr(),
            h4("How to proceed:"),
      h5("1- Upload the desired proposals to convert"),
      h5("2- Select desired default parameters (questions/review)"),
      h5("3- Click to summarize"),
      h5("4- Result table of summarized proposal shows if they were a success"),
      h5("5- Download!"),
      hr(),
      "Source code can be found here:"
     # tags$i(class="fa fa-github fa-2x", style="color:#FAFAFA")
      
    ),

    mainPanel(
      shinyjs::useShinyjs(),
      fluidRow(
        column(width = 6, offset = 1,
          fileInput("file", label = "Upload your proposal(s)",
                     placeholder = "or drop files here",
                     multiple = TRUE,
                     accept = c(".xlsx", ".xlsm", ".xls"),
                     width = '100%')
        ),
      ), 

      br(),

      fluidRow(
        column(width = 4, offset = 1,
          radioButtons("questions", "Including questions to expert",
                      choices = list("Yes" = "withquestion",
                                 "No" = "withoutquestion"),
                      selected = "withoutquestion"),
        ),
        column(width = 4,
        # radioButtons to select if you want to include reviews
          radioButtons("reviews", "Including review",
                        choices = list("Yes" = "withreview",
                                   "No" = "withoutreview"),
                        selected = "withoutreview"),
        ),
      ),

      fluidRow(
        column(width = 2, offset = 1, uiOutput('runButton')),
        column(width = 2, offset = 1, downloadButton("downloadData", "Download"))
      ),

      br(),
      
      fluidRow(
        column(width = 8, offset = 1, tableOutput("resTable")),
      )
    )
  )
)

server <- function(input, output) {

source("R/KBA_summary.R")

shinyjs::hide('downloadData')

  file_df <- reactive({
    req(input$file)
    df <- input$file
  })

  output$runButton <- renderUI({
    if(is.null(file_df())) return()
    actionButton("runScript", "Convert to summary")
  })

  askQuestion <- reactive({
    if(input$questions == "withquestion") return(TRUE)
    if(input$questions == "withoutquestion") return(FALSE)
  })

  askReview <- reactive({
    if(input$reviews == "withreview") return(TRUE)
    if(input$reviews == "withoutreview") return(FALSE)
  })

  r <- reactiveValues(convertRes = NULL)

  observeEvent(input$runScript, {
    shinyjs::disable("runScript")
    r$convertRes <- form_conversion(KBAforms = file_df()$datapath, includeQuestions = askQuestion(), includeReviewDetails = askReview())
    
    output$resTable <- renderTable(r$convertRes[[2]])

    output$downloadData <- downloadHandler(
      filename = function() if(length(file_df()$name) == 1){
        paste0(r$convertRes[[2]]$Name,".docx")
      } else{"Summaries.zip"},
      content = function(file) if(length(file_df()$name) == 1){
        file.rename( from = r$convertRes[[1]], to = file )
      } else{
          # create a temp folder for shp files
          temp_fold <- tempdir()
          zip_file <- paste0(temp_fold,"/Summaries.zip")
          zip(zipfile = zip_file, files = r$convertRes[[1]])
          # copy the zip file to the file argument
          file.copy(zip_file, file, overwrite = TRUE)
          # remove all the files created
          file.remove(zip_file)
        }
    
    )

    rm(delineationRationale,includeGlobalTriggers,includeNationalTriggers,juris,lat,lon,nationalName,proposalLead,scope,siteDescription, envir = sys.frame())
    shinyjs::show('downloadData')

  }, ignoreNULL = TRUE, ignoreInit = TRUE)
  
}

shinyApp(ui, server)
