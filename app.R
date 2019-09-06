library(shiny)
library(tidyverse)
library(janitor)
library(scales)
library(lubridate)


source("data/locations.R")

# based on the Shiny fileInput function
fileInput2 <- function(inputId, label = NULL, labelIcon = NULL, multiple = FALSE, 
                       accept = NULL, width = NULL, progress = TRUE, ...) {
  # add class fileinput_2 defined in UI to hide the inputTag
  inputTag <- tags$input(id = inputId, name = inputId, type = "file", 
                         class = "fileinput_2")
  if (multiple) 
    inputTag$attribs$multiple <- "multiple"
  if (length(accept) > 0) 
    inputTag$attribs$accept <- paste(accept, collapse = ",")
  
  div(..., style = if (!is.null(width)) paste0("width: ", validateCssUnit(width), ";"), 
      inputTag,
      # label customized with an action button
      tags$label(`for` = inputId, div(icon(labelIcon), label, 
                                      class = "btn btn-default action-button")),
      # optionally display a progress bar
      if(progress)
        tags$div(id = paste(inputId, "_progress", sep = ""), 
                 class = "progress shiny-file-input-progress", 
                 tags$div(class = "progress-bar")
        )
  )
}          


ui <- fluidPage(
  
  # App title ----
  titlePanel("Upload the specified files and select your dates, then press download"),
  
  # define class fileinput_2 to hide inputTag in fileInput2
  tags$head(tags$style(HTML(
    ".fileinput_2 {
      width: 0.1px;
      height: 0.1px;
      opacity: 0;
      overflow: hidden;
      position: absolute;
      z-index: -1;
    }"
  ))),
  sidebarLayout(
    sidebarPanel(
      h3("Attach JJ Keller CSV"),  
      fileInput2("file1", "Load File 1", labelIcon = "folder-open-o", 
                 accept = c("text/csv",
                            "text/comma-separated-values,text/plain",
                            ".csv"), progress = TRUE),
      h3("Attach CBCS Loss Run Report CSV"),
      fileInput2("file2", "Load File 2", labelIcon = "folder-open-o", 
                 accept = c("text/csv",
                            "text/comma-separated-values,text/plain",
                            ".csv"), progress = TRUE),
      # Date range for CBCS Desc ----
      dateRangeInput('dateRange',
                     label = 'Date range input: yyyy-mm-dd',
                     start = Sys.Date() - 6, end = Sys.Date()),
                     
      # Territory selection ----
      selectInput("select", label = h3("Territory"), 
                  choices = list("TAMPA" = "TAMPA", "PLACEHOLDER" = "PLACEHOLDER"))
      ),
    mainPanel(
      
      tabsetPanel(type = "tabs",
                  tabPanel("Info",  verbatimTextOutput("content1"),
                                    verbatimTextOutput("content2"),
                                    verbatimTextOutput("dateRangeText")),
                  tabPanel("JJ Keller",  plotOutput(outputId = "plots")))
     
    )
  )
)

server <- function(input, output, session) {
  
  df <- reactive({
    req(input$file1)
    read.csv(input$file1$datapath, header = T) %>%
      filter(DQ.File == "In Compliance" | DQ.File == "Out of Compliance") %>% 
      group_by(DQ.File) %>% 
      summarise (n = n()) %>% 
      mutate(DQ.File = recode(DQ.File,
                              "In Compliance" = "Drivers In Compliance",
                              "Out of Compliance" = "Drivers Out Of Compliance"),
             freq = round((n / sum(n)) * 100, 2),
             label = paste(DQ.File, "-", paste(freq, "%", sep = ""))) %>% 
      select(-c(n, DQ.File))
  })
  
  output$plots <- renderPlot({
    df() %>% 
      ggplot(aes(x = 1, y = freq, fill = label)) +
      coord_polar(theta = 'y') +
      geom_bar(stat = "identity", color = 'black') +
      scale_fill_manual(values = c("darkgreen", "red")) +
      theme_minimal()+
      theme(
        axis.title.x = element_blank(),
        axis.text = element_blank(),
        axis.title.y = element_blank(),
        panel.border = element_blank(),
        panel.grid = element_blank(),
        axis.ticks = element_blank(),
        plot.title = element_text(size=14, face="bold"),
        legend.title = element_blank(),
        axis.text.x = element_blank(),
        legend.background = element_rect(linetype = "solid"))
  })
}

shinyApp(ui, server)
