# Needed libraries to run the app
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
library(shinyWidgets)
library(readxl)
library(readr)
library(devtools)

# Authenticate googledrive
# Informations to authenticate are stored in an .Rprofile
googledrive::drive_auth()

# User interface
ui <- fluidPage(

# Calling of shinyjs package for functionality in `server`
  shinyjs::useShinyjs(),
  theme = bs_theme(bg = "#f6f6f6", fg = "#1D4044",
  base_font = font_google("Lato"), "progress-bar-bg" = "orange"),

# Change web window title + adding the CSS styling 
  tags$head(
    HTML("<title> Create KBA summary </title>"),
    includeCSS("www/styles.css")
  ),

# Add the titlepanel which is modified in the .css file
  HTML("<h1><a href='https://kbacanadawiki.org/'>
  <img src='./Canada_KBA_transparent_white.png'style='position: relative;
  float: left;left: 6%; padding-bottom: 10px;
   width: 5%; height: 5%;' alt='KBA_logo'/></a>Creation of KBA summaries</h1>"),

# Spaces before the panels
  br(),

  fluidRow(
    column(width = 3, offset = 1,
      wellPanel(class = "well", h2("ReadMe"),
        hr(),
        h5("This Shiny application is intended to convert KBA
         Canada proposal forms into KBA summaries for expert review."),
        h3("How to proceed:"),
        hr(),
        h4("1- Select review stage."),
        h4("2- Select language."),
        h4("3- Upload the desired proposals to convert."),
        h4("4- Click convert button to summarize."),
        h4("(Process might take a couple seconds to start)"),
        h4("5- Check result table to see which
         proposals were correctly processed."),
        h4("6- Download!"),
        hr(),
        h5("Developed by Benjamin Mercier and
         Chloé Debyser for the KBA Canada Secretariat"),
        h6("Source code",
         tags$a(href = "https://github.com/BenMerSci/KBA_shinyforms",
         icon("github", "fa-2x"))),
        tags$style(".fa-github {color:#13294B}"),
      )
    ),

    column(width = 4,
      wellPanel(class = "well", h2("Input"),
        hr(),
        
        radioButtons(
          "stageRev",
          h3("Select review stage"),
          choices = list("Technical Review" = "technicalRev",
                         "General Review" = "generalRev",
                         "Steering Committee" = "steeringRev"),
          selected = "technicalRev", inline = TRUE),
        
        tags$div(
          h3("Select language"),
          materialSwitch(
            inputId = "language",
            label = h4("English",
                       style = "margin-top: -15%; display: inline"),
            value = F,
            inline = T,
            right = F,
            width="50%"),
          tags$span(h4("Français",
                       style = "margin-left: -28%; display: inline"))
        ),

          fileInput("file", label = h3("Upload your proposal(s)"),
           placeholder = "or drop files here",
           multiple = TRUE,
           accept = c(".xlsm"),
           width = "120%"),

        hr(),
        br(),

          tags$div(uiOutput("runButton"), align = "center"),
      )
    ),

    column(width = 4,
      wellPanel(class = "well_scroll", h2("Output"),
       hr(),
       br(),
        tags$div(downloadButton("downloadData", "Download"), align = "center"),
       br(),
        tags$div(tableOutput("resTable"), align = "center"),
      )
    ),


  ),

  fluidRow(
    column(12,
       tags$div(class = "footer",
         HTML("Photograph © Tony Webster, <a rel='noreferrer noopener'
         href='https://www.flickr.com/photos/diversey/22492133101/in/photolist-
         Agy4yi-4hP5D-2jiQL9m-2iqMZ89-Q4p8C3-gAnYRA-24Z34E4-PuPSu3-2kjavbz-
         2kbsmz1-2iqP7Fu-nmFr4C-zbb2AM-qeoBfK-nYdH9B-RASxd7-Qd1Sgq-2f2oPXZ-
         qzAw71-8MVKVd-T7niUj-etMN7D-DkStuM-55gdqb-2j8xytN-qcnKee-gAo3X9-ur89y7-
         gAkQjM-5iBJJR-diC8TQ-aS2LbM-pq2Y4k-5iG1qG-5iBJiZ-g88dhq-2j8xywt-5iBJh8-
         u9XuRf-PE5GWE-2kJ5z3N-5iG1SY-JTZ92V-5iBJBn-3gaoc2-Cudqzt-2koJsqF-
         2jd9VDJ-6y7wxz-dk4Aui/' target='_blank'>Lake of Shining Waters -
         Prince Edward Island National Park</a>
         (cropped) <a rel='noreferrer noopener'
         href='https://creativecommons.org/licenses/by-sa/2.0/'
         target='_blank'>CC BY-SA 2.0</a>")
        )
    )
  )
)

server <- function(input, output) {

source("R/KBA_summary.R")


shinyjs::hide("downloadData")

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
    if(input$stageRev == "steeringRev") return("steering")
  })
  
  getLanguage <- reactive({
    if(!input$language) return("english")
    if(input$language) return("french")
  })

  r <- reactiveValues(convertRes = NULL)

  observeEvent(input$runScript, {
    
    shinyjs::disable("runScript")
    shinyjs::hide("downloadData")
    shinyjs::hide("resTable")
    
    r$convertRes <- form_conversion(KBAforms = file_df()$datapath, reviewStage = getReviewStage(), language = getLanguage(), app = T)

    output$resTable <- renderTable(r$convertRes[[2]])
    shinyjs::show("resTable")
    
    output$downloadData <- downloadHandler(
      filename = function() if(length(file_df()$name) == 1){
        r$convertRes[[1]]
      }else{"Summaries.zip"},
      content = function(file) if(length(file_df()$name) == 1){
        file.rename(from = r$convertRes[[1]], to = file)
      }else{
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

    rm(delineationRationale, includeGlobalTriggers, includeNationalTriggers,
    juris, lat, lon, nationalName, proposalLead, scope, siteDescription,
    noFeedback, envir = sys.frame())
    shinyjs::show("downloadData")

  }, ignoreNULL = TRUE, ignoreInit = TRUE)

}


shinyApp(ui, server)
