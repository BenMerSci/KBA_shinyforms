library(shiny)
library(tidyverse)
library(magrittr)
library(WordR)
library(flextable)
library(officer)
library(openxlsx)
library(lubridate)
library(janitor)
library(googledrive)

# Authenticate googledrive
# Informations to authenticate are stored in an .Rprofile
googledrive::drive_auth()

# User interface
ui <- fluidPage(
  sidebarLayout(
    sidebarPanel(      
        h1("Instructions")
    ),

    mainPanel(
      titlePanel("KBA proposals conversion"),
      shinyjs::useShinyjs(),
        # radioButtons to select weither you want to summarize one or multiple proposals
      fluidRow(
        column(width = 4,
          radioButtons("proposalNumber", "Number of proposals to summarize",
                      choices = list("One proposal" = "1prop", 
                                  "Multiple proposals" = "xprop"),
                      selected = "1prop"),
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
      # input fields change depending on proposalNumber wanted
        conditionalPanel(
          condition = "input.proposalNumber == '1prop'",
          fileInput("file1", "Select proposal to summarize",
                     # select just selected sheet
                     multiple = FALSE,
                     accept = c(".xlsx", ".xlsm", ".xls"),
                     width = '100%')
        ),
      
        conditionalPanel(
          condition = "input.proposalNumber == 'xprop'",
          fileInput("file2", "Select proposals to summarize",
                     multiple = TRUE,
                     accept = c(".xlsx", ".xlsm", ".xls"),
                     width = '100%')
        ),

      
        actionButton("runScript", "Convert to summary"),
        downloadButton("downloadData", "Download")
    )
  )
)

server <- function(input, output) {

source("R/KBA_summary.R")

shinyjs::hide('downloadData')

  file_df <- reactive({
    req(input$proposalNumber)
    
    if (input$proposalNumber == "1prop") {
      req(input$file1)
      df <- input$file1
    }else if (input$proposalNumber == "xprop") {
      req(input$file2)
      df <- input$file2
    }
    
  })

  askQuestion <- reactive({
    if(input$questions == "withquestion") return(TRUE)
    if(input$questions == "withoutquestion") return(FALSE)
  })

  askReview <- reactive({
    if(input$reviews == "withreview") return(TRUE)
    if(input$reviews == "withoutreview") return(FALSE)
  })

  r <- reactiveValues(test = NULL)

  observeEvent(input$runScript, {
    r$test <- form_conversion(KBAforms = file_df()$datapath, includeQuestions = askQuestion(), includeReviewDetails = askReview())
    
    output$downloadData <- downloadHandler(
      filename = function() "Summaries.zip",
      content = function(file) {
          # create a temp folder for shp files
          temp_fold <- tempdir()
          zip_file <- paste0(temp_fold,"/Summaries.zip")
          zip(zipfile = zip_file, files = r$test)
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
