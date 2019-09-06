library(shiny)
library(tidyverse)
library(janitor)
library(scales)
library(lubridate)


source("data/locations.R")

# Define UI for data upload app ----
ui <- fluidPage(
  
  # App title ----
  titlePanel("Uploading Files"),
  
  # Sidebar layout with input and output definitions ----
  sidebarLayout(
    
    # Sidebar panel for inputs ----
    sidebarPanel(
      
      # Input: Select a file ----
      fileInput("file1", "Choose CSV File",
                multiple = TRUE,
                accept = c("text/csv",
                           "text/comma-separated-values,text/plain",
                           ".csv"))
      
    ),
    
    # Main panel for displaying outputs ----
    mainPanel(
      
      # Output: Data file ----
      plotOutput(outputId = "plots")
      
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
