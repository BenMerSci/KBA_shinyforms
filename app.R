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

# Create KBA summaries
# User interface
ui <- fluidPage(
shinyjs::useShinyjs(),
theme = bs_theme("progress-bar-bg" = "orange", base_font = font_google("Lato")),

titlePanel(
fluidRow(
    column(11, h1("Creation of KBA summaries")),
    column(1, img(src = "Canada_KBA_transparent.png", height = 80, width = 105, href = "https://www.kbacanadawiki.org"))
    
  )
),

  sidebarLayout(

    sidebarPanel(id = "sidebar", width = 4,  
      h3("ReadMe", align = "center"),
      "This Shiny application serves to convert KBA Canada proposal forms into KBA summaries for expert review.",
      hr(),
            h4("How to proceed:"),
      h5("1- Upload the desired proposals to convert"),
      h5("2- Select review stage"),
      h5("3- Click convert button to summarize"),
      h5("4- Check result table to see which proposals were correctly processed."),
      h5("5- Download!"),
      hr(),
      h6("Developed by Benjamin Mercier and Chloé Debyser for the KBA Canada Secretariat"),
      h6("Source code", tags$a(href="https://github.com/BenMerSci/KBA_shinyforms", icon("github","fa-2x"))),
      tags$style(".fa-github {color:#13294B}")
    ),

    mainPanel(
      fluidRow(
        column(width = 6, offset = 1,
          fileInput("file", label = "Upload your proposal(s)",
                     placeholder = "or drop files here",
                     multiple = TRUE,
                     accept = c(".xlsx", ".xlsm", ".xls"),
                     width = '100%')
        )  
      ),

      br(),

      fluidRow(
        column(width = 4, offset = 1,
        # radioButtons to select which stage of review
          radioButtons("stageRev", "Select review stage",
                        choices = list("Technical review" = "technicalRev",
                                   "General review" = "generalRev",
                                   "Steering Committee" = "steeringRev"),
                        selected = "technicalRev"),
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

  getReviewStage <- reactive({
    if(input$stageRev == "technicalRev") return("technical")
    if(input$stageRev == "generalRev") return("general")
    if(input$stageRev == "steergingRev") return("steering")
  })

  r <- reactiveValues(convertRes = NULL)

  observeEvent(input$runScript, {
    shinyjs::disable("runScript")
    r$convertRes <- form_conversion(KBAforms = file_df()$datapath, reviewStage = getReviewStage())
    
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

    rm(delineationRationale,includeGlobalTriggers,includeNationalTriggers,juris,lat,lon,nationalName,proposalLead,scope,siteDescription,noFeedback, envir = sys.frame())
    shinyjs::show('downloadData')

  }, ignoreNULL = TRUE, ignoreInit = TRUE)
  
}

shinyApp(ui, server)
