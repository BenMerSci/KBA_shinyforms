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
  theme = bs_theme(version = 5),
  sidebarLayout(

    sidebarPanel(      
        h3("README")
    ),

    mainPanel(
      titlePanel("Creation of KBA summaries"),
      shinyjs::useShinyjs(),

      fluidRow(
        column(width = 4,
          fileInput("file", label = "Upload your proposal(s)",
                     placeholder = "or drop files here",
                     multiple = TRUE,
                     accept = c(".xlsx", ".xlsm", ".xls"),
                     width = '100%')
        ),
        # radioButtons to select if you want to include question to experts
        column(width = 4,
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

        tableOutput("resTable"),
        uiOutput('runButton'),
        downloadButton("downloadData", "Download")
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
