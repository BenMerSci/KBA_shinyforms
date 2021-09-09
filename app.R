library(shiny)
library(xlsx)
# User interface
ui <- fluidPage(

  fileInput("uploadFile", "files_xlsx", multiple = TRUE)
  #submitButton(text = "Sumbit")

)

# Server side
server <- function(input, ouput) {}

# Run the application
shinyApp(ui = ui, server = server)
