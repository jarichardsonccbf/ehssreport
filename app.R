library(shiny)
library(tidyverse)
library(janitor)
library(scales)
library(lubridate)


source("data/locations.R")

# Function for uploading multiple. I have no idea how this works.
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
  titlePanel("Upload the files indicated and select your dates and territory"),
  
  # define class fileinput_2 to hide inputTag in fileInput2. Not sure what this is doing.
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
  
  # Side bar layout and inputs ----
  sidebarLayout(
    sidebarPanel(
      
      # File 1 upload ----
      h3("Attach JJ Keller CSV"),  
      fileInput2("file1", "Load File 1", labelIcon = "folder-open-o", 
                 accept = c("text/csv",
                            "text/comma-separated-values,text/plain",
                            ".csv"), progress = TRUE),
      
      # File 2 upload ----
      h3("Attach CBCS Loss Run Report CSV"),
      fileInput2("file2", "Load File 2", labelIcon = "folder-open-o", 
                 accept = c("text/csv",
                            "text/comma-separated-values,text/plain",
                            ".csv"), progress = TRUE),
      
      # Date range for CBCS Desc ----
      dateRangeInput('dateRange',
                     label = 'Input date range',
                     start = Sys.Date() - 6, end = Sys.Date()),
      
      # Territory selection ----
      selectInput("test.selection", label = h3("Territory"), 
                  choices = list("TAMPA", "PLACEHOLDER"))
      
      # Another territory selection?
    ),
    
    # Main panel display. Use tabs, one for each slide ----
    mainPanel(
      
      # Set tabs ----
      tabsetPanel(type = "tabs",
                  # JJ Keller infos tab
                  tabPanel("JJ Keller",  plotOutput(outputId = "jjk.compliance.pie"),
                           plotOutput(outputId = "jjk.driver.stats.pie"),
                           tableOutput(outputId = "driver.qual.table")),
                  # Incident descriptions tab
                  tabPanel("Weekly Safety Incidents", tableOutput(outputId = "weekly.cbcs.incidents")),
                  # Costs and Counts tab
                  tabPanel("Claims Costs and Counts", tableOutput(outputId = "cost.pivot"),
                           tableOutput(outputId = "count.pivot"))
      )
      
    )
  )
)

server <- function(input, output, session) {
  
  # JJ Keller Functions ----
  
  # jjk compliance pie df
  jjk.compliance.pie.df <- reactive({
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
  
  # jjk compliance pie plot
  output$jjk.compliance.pie <- renderPlot({
    jjk.compliance.pie.df() %>% 
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
  
  # jjk driver stats pie df
  jjk.driver.stats.pie.df <- reactive({
    req(input$file1)
    read.csv(input$file1$datapath, header = T) %>%
      group_by(Status) %>%
      summarise (n = n()) %>%
      mutate(freq = round((n / sum(n)) * 100, 2),
             Status = recode(Status,
                             "Not Driving-" = "Not Driving"),
             label = paste(Status, "-", paste(freq, "%", sep = ""))) %>%
      select(-c(n, Status))  
  })
  
  # jjk driver stats pie plot
  output$jjk.driver.stats.pie <- renderPlot({
    jjk.driver.stats.pie.df() %>% 
      ggplot(aes(x = 1, y = freq, fill = label)) +
      coord_polar(theta='y') +
      geom_bar(stat = "identity", color = 'black') +
      scale_fill_manual(values = c("deepskyblue4",
                                   "firebrick4",
                                   "yellowgreen",
                                   "darkslateblue",
                                   "darkcyan",
                                   "chocolate3",
                                   "lightsteelblue2",
                                   "lightpink4",
                                   "darkolivegreen4")) +
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
        legend.background = element_rect(linetype = "solid")) +
      guides(fill = guide_legend(override.aes = list(colour=NA)))
    
  })
  
  # driver qualification tables
  jjk.qual.table <- reactive({
    req(input$file1)
    jjk.qual <- read.csv(input$file1$datapath, header = T) %>% 
      left_join(jjkeller.locations, "Assigned.Location") %>% 
      filter(manager == input$test.selection, 
             DQ.File == "In Compliance" | DQ.File == "Out of Compliance") %>%
      mutate(dq.binary = recode(DQ.File,
                                "In Compliance" = 1,
                                "Out of Compliance" = 0)) %>% 
      group_by(Assigned.Location) %>% 
      summarise(num.compliant     = sum(dq.binary      ),
                total             = length(dq.binary   ),
                percent.compliant = num.compliant/total * 100) 
    
    jjk.qual.totals <- data.frame("Total:", sum(jjk.qual$num.compliant), sum(jjk.qual$total), round(sum(jjk.qual$num.compliant)/sum(jjk.qual$total) * 100,2))
    
    colnames(jjk.qual.totals) <- colnames(jjk.qual)
    
    # driver qual table
    rbind(jjk.qual,jjk.qual.totals)
  })
  
  output$driver.qual.table <- renderTable({
    jjk.qual.table()
  })
  
  # CBCS details ----
  
  # Incidents
  weekly.incidents <- reactive({
    req(input$file2)
    read.csv(input$file2$datapath, header = T) %>% 
      # rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
      left_join(cbcs.locations, by = "Dept.Name") %>% 
      mutate(Coverage = recode(Coverage,
                               "ALBI"  = "AUTO",
                               "ALPD"  = "AUTO",
                               "ALAPD" = "AUTO",
                               "GLPD"  = "GL"  ,
                               "GLBI"  = "GL")) %>%
      filter(manager == input$test.selection) %>%
      filter(as.Date(Occ.Date, format = "%m/%d/%Y") >= input$dateRange[1] & as.Date(Occ.Date, format = "%m/%d/%Y") <= input$dateRange[2]) %>% # change to `Occ Date` if using Xls
      select(manager, Dept.Name, Occ.Date, Coverage, Acc.Desc)
  })
  
  output$weekly.cbcs.incidents <- renderTable({
    weekly.incidents()
  })
  
  # Pivots
  cbcs.pivots.df <- reactive({
    req(input$file2)
    read.csv(input$file2$datapath, header = T) %>% 
      left_join(cbcs.locations, by = "Dept.Name") %>% 
      mutate(Coverage = recode(Coverage,
                               "ALBI"  = "AUTO",
                               "ALPD"  = "AUTO",
                               "ALAPD" = "AUTO",
                               "GLPD"  = "GL"  ,
                               "GLBI"  = "GL"),
             occ.year = year(as.Date(Occ.Date, format = "%m/%d/%Y"))) %>% 
      filter(manager == input$test.selection,
             occ.year == format(Sys.Date(), "%Y"))
  })
  
  # Cost Pivot
  output$cost.pivot <- renderTable({
    costs <- cbcs.pivots.df() %>% 
      group_by(Loc.Name, Coverage) %>% 
      mutate(raw = as.numeric(gsub('[$,]', '', Incurred))) %>% 
      summarise(Incurred = sum(raw)) %>%
      ungroup() %>% 
      spread(Coverage, Incurred) %>%
      replace(., is.na(.), 0) %>% 
      mutate(Totals = rowSums(.[2:4]))
    
    costs.tot <- data.frame("Total", sum(costs$AUTO), sum(costs$GL), sum(costs$WC), sum(costs$Totals))
    
    colnames(costs.tot) <- colnames(costs)
    
    # claim cost table
    rbind(costs, costs.tot)
  })
  
  # Count Pivot
  output$count.pivot <- renderTable({
    counts <- cbcs.pivots.df() %>% 
      group_by(Loc.Name, Coverage) %>% 
      summarise (n = n()) %>% 
      spread(Coverage, n) %>% 
      replace(., is.na(.), 0) %>% 
      mutate(Totals = sum(AUTO, GL, WC)) %>% 
      ungroup()
    
    counts.tot <- data.frame("Total", sum(counts$AUTO), sum(counts$GL), sum(counts$WC), sum(counts$Totals))
    
    colnames(counts.tot) <- colnames(counts)
    
    # claim count table
    claim.count <- rbind(counts,counts.tot)
  })
}

shinyApp(ui, server)
