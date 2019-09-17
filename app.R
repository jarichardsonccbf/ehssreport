library(shiny)
library(tidyverse)
library(janitor)
library(scales)
library(lubridate)
library(readxl)
library(flextable)
library(officer)

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
  titlePanel("Upload the files indicated and select your dates, then press download"),
  
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
      
      # JJK upload ----
      h3("Attach JJ Keller"),  
      fileInput2("file1", "File location", labelIcon = "folder-open-o", 
                 accept = c("text/csv",
                            "text/comma-separated-values,text/plain",
                            ".csv"), progress = TRUE),
      
      # CBCS upload ----
      h3("Attach CBCS Loss Run Report"),
      fileInput2("file2", "File location", labelIcon = "folder-open-o", 
                 accept = c(".xlsx"), progress = TRUE),
      
      # LMS upload ----
      h3("Attach LMS Data Pull"),
      fileInput2("file3", "File location", labelIcon = "folder-open-o", 
                 accept = c(".csv"), progress = TRUE),
      
      # All territories or not ----
      selectInput("allstate", label = h3("Entire territory?"),
                  choices = list("NO", "YES")),
      
      # Territory selection ----
      selectInput("manager", label = h3("Region"), 
                  choices = list("TAMPA", "PLACEHOLDER")),
      
      # Date range for CBCS Desc ----
      dateRangeInput('dateRange',
                     label = h3("Input date range for weekly incident table"),
                     start = Sys.Date() - 6, end = Sys.Date()),
      
      # Download buttom ----
      downloadButton("download_powerpoint", "Download Tables to Powerpoint")
      
    ),
    
    # Main panel display. Use tabs, one for each slide ----
    mainPanel(
      
      # Set tabs ----
      tabsetPanel(type = "tabs",
                  # General intro tab
                  tabPanel("Intro",
                           h3(p("Download data from JJ Keller, Successfactors, LMS, CBCS and STARS.", 
                                span("Do not", style = "color:red"),
                                "change any information in your raw data pulls. Upload the indicated data sources to the left.")),
                           h3(p("Please report any changes in reporting vendors, or territory changes to Jason Richardson (",
                                span("jarichardson@cocacolaflorida.com", style = "color:blue"),
                                ")."))),
                  # JJ Keller infos tab
                  tabPanel("JJ Keller",  plotOutput(outputId = "jjk.compliance.pie"),
                           plotOutput(outputId = "jjk.driver.stats.pie"),
                           tableOutput(outputId = "driver.qual.table")),
                  # Incident descriptions tab
                  tabPanel("Weekly Safety Incidents", tableOutput(outputId = "weekly.cbcs.incidents")),
                  # Costs and Counts tab
                  tabPanel("Claims Costs and Counts", tableOutput(outputId = "cost.pivot"),
                           tableOutput(outputId = "count.pivot")),
                  #LMS tab
                  tabPanel("LMS", tableOutput(outputId = "lms.pivot"))
      )
      
    )
  )
)

server <- function(input, output, session) {
  
  # JJK compliance pie chart ----
  
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
  
  # JJK driver states pie chart ----
  
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
  
  # JJK driver qual table ---- 
  
  # driver qualification tables
  jjk.qual.table <- reactive({
    req(input$file1)
    
    if(input$allstate == "NO")
      
      jjk.qual <- read.csv(input$file1$datapath, header = T) %>% 
        left_join(jjkeller.locations, "Assigned.Location") %>% 
        filter(manager == input$manager, 
               DQ.File == "In Compliance" | DQ.File == "Out of Compliance")
    
    else
      
      jjk.qual <- read.csv(input$file1$datapath, header = T) %>% 
        left_join(jjkeller.locations, "Assigned.Location") %>% 
        filter(DQ.File == "In Compliance" | DQ.File == "Out of Compliance")
    
    jjk.qual <- jjk.qual %>% 
      mutate(dq.binary = recode(DQ.File,
                                "In Compliance" = 1,
                                "Out of Compliance" = 0)) %>% 
      group_by(Assigned.Location) %>% 
      summarise(num.compliant     = sum(dq.binary      ),
                total             = length(dq.binary   ),
                percent.compliant = num.compliant/total * 100) 
    
    jjk.qual.totals <- data.frame("Total", sum(jjk.qual$num.compliant), sum(jjk.qual$total), round(sum(jjk.qual$num.compliant)/sum(jjk.qual$total) * 100,2))
    
    colnames(jjk.qual.totals) <- colnames(jjk.qual)
    
    # driver qual table
    rbind(jjk.qual,jjk.qual.totals)
  })
  
  output$driver.qual.table <- renderTable({
    jjk.qual.table()
  })
  
  # CBCS weekly incidents ----
  
  # Incidents
  weekly.incidents <- reactive({
    req(input$file2)
    
    if(input$allstate == "YES")
      
      cbcs.state <- read_excel(input$file2$datapath, skip = 6)  %>% 
        rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        left_join(cbcs.locations, by = "Dept.Name") 
    
    
    else
      
      cbcs.state <- read_excel(input$file2$datapath, skip = 6) %>% 
        rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        left_join(cbcs.locations, by = "Dept.Name") %>% 
        filter(manager == input$manager)
    
    cbcs.state %>% 
      mutate(Coverage = recode(Coverage,
                               "ALBI"  = "AUTO",
                               "ALPD"  = "AUTO",
                               "ALAPD" = "AUTO",
                               "GLPD"  = "GL"  ,
                               "GLBI"  = "GL") ,
             Incident = paste(`Occ Date`, "-",Coverage, "-",`Acc Desc`)) %>% 
      filter(`Occ Date` >= input$dateRange[1] & `Occ Date` <= input$dateRange[2]) %>%   # change to `Occ Date` if using Xls
      select(Incident)
  })
  
  output$weekly.cbcs.incidents <- renderTable({
    weekly.incidents()
  })
  
  # DF for CBCS pivots ----
  
  cbcs.pivots.df <- reactive({
    req(input$file2)
    
    if(input$allstate == "YES")
      
      cbcs.pivots <- read_excel(input$file2$datapath, skip = 6) %>% 
        rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        left_join(cbcs.locations, by = "Dept.Name") 
    
    else
      
      cbcs.pivots <- read_excel(input$file2$datapath, skip = 6) %>% 
        rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        left_join(cbcs.locations, by = "Dept.Name") %>% 
        filter(manager == input$manager)
    
    cbcs.pivots %>%
      mutate(Coverage = recode(Coverage,
                               "ALBI"  = "AUTO",
                               "ALPD"  = "AUTO",
                               "ALAPD" = "AUTO",
                               "GLPD"  = "GL"  ,
                               "GLBI"  = "GL"),
             occ.year = year(as.Date(`Occ Date`, format = "%m/%d/%Y"))) %>% 
      filter(occ.year == format(Sys.Date(), "%Y"))
  })
  
  # Cost Pivot ----
  
  cost.pivot.flex <- reactive({
    costs <- cbcs.pivots.df() %>% 
      group_by(`Loc Name`, Coverage) %>% 
      mutate(raw = as.numeric(gsub('[$,]', '', Incurred))) %>% 
      summarise(Incurred = sum(raw)) %>%
      ungroup() %>% 
      spread(Coverage, Incurred) %>%
      replace(., is.na(.), 0) %>% 
      mutate(Totals = rowSums(.[2:4]))
    
    costs.tot <- data.frame("Total", sum(costs$AUTO), sum(costs$GL), sum(costs$WC), sum(costs$Totals))
    
    colnames(costs.tot) <- colnames(costs)
    
    # claim cost table
    claim.cost <- rbind(costs, costs.tot)
  })
  
  output$cost.pivot <- renderTable({
    costs <- cbcs.pivots.df() %>% 
      group_by(`Loc Name`, Coverage) %>% 
      mutate(raw = as.numeric(gsub('[$,]', '', Incurred))) %>% 
      summarise(Incurred = sum(raw)) %>%
      ungroup() %>% 
      spread(Coverage, Incurred) %>%
      replace(., is.na(.), 0) %>% 
      mutate(Totals = rowSums(.[2:4]))
    
    costs.tot <- data.frame("Total", sum(costs$AUTO), sum(costs$GL), sum(costs$WC), sum(costs$Totals))
    
    colnames(costs.tot) <- colnames(costs)
    
    # claim cost table
    claim.cost <- rbind(costs, costs.tot)
  })
  
  # Count Pivot----
  
  count.pivot.flex <- reactive({
    counts <- cbcs.pivots.df() %>% 
      group_by(`Loc Name`, Coverage) %>% 
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
  
  output$count.pivot <- renderTable({
    counts <- cbcs.pivots.df() %>% 
      group_by(`Loc Name`, Coverage) %>% 
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
  
  # LMS Pivot ----
  lms.pivots.df <- reactive({
    
    req(input$file3)
    
    if(input$allstate == "YES")
      
      lms <- read.csv(input$file3$datapath) %>% 
        left_join(lms.locations, "Org.Name")
    
    else
      
      lms <- read.csv("data/lms.csv") %>% 
        left_join(lms.locations, "Org.Name") %>%
        filter(manager == input$manager)
    
    lms %>% mutate(Item.Status = recode(Item.Status,
                                        "In Progress" = "Incomplete",
                                        "Not Started" = "Incomplete"),
                   comp.binary = recode(Item.Status,
                                        "Incomplete" = 0,
                                        "Completed" = 1)) %>%
      group_by(location, Title) %>%
      summarise(num.comp = sum(comp.binary),
                total = length(comp.binary),
                percent.compliant = num.comp/total * 100) %>% 
      select(location, Title, percent.compliant) %>% 
      pivot_wider(names_from = Title, values_from = percent.compliant)
    
  })
  
  output$lms.pivot <- renderTable({
    lms.pivots.df()
  })   
  
  # PPT
  
  output$download_powerpoint <- downloadHandler(
    filename = function() {  
      "EHSStest.pptx"
    },
    content = function(file) {
      
      # jjk flex ----
      flextable_dq <- flextable(jjk.qual.table()) %>% 
        # colformat_num(col_keys = c("Location", "# Compliant Employees", "Total Employees", "Percent Compliant"), digits = 0) %>% 
        width(width = 1.25) %>% 
        height_all(height = 0.35) %>% 
        theme_zebra() %>% 
        align(align = "center", part = "all")
      
      # cbcs incidents flex ----
      flextable_cbcs.inc <- flextable(weekly.incidents()) %>% 
        # colformat_num(col_keys = c("Location", "# Compliant Employees", "Total Employees", "Percent Compliant"), digits = 0) %>% 
        width(width = 1.25) %>% 
        height_all(height = 0.35) %>% 
        theme_zebra() %>% 
        align(align = "center", part = "all")
      
      # cbcs cost flex ----
      flextable_cbcs.cost <- flextable(cost.pivot.flex()) %>% 
        # colformat_num(col_keys = c("Location", "# Compliant Employees", "Total Employees", "Percent Compliant"), digits = 0) %>% 
        width(width = 1.25) %>% 
        height_all(height = 0.35) %>% 
        theme_zebra() %>% 
        align(align = "center", part = "all")
      
      # cbcs count flex ----
      flextable_cbcs.count <- flextable(count.pivot.flex()) %>% 
        # colformat_num(col_keys = c("Location", "# Compliant Employees", "Total Employees", "Percent Compliant"), digits = 0) %>% 
        width(width = 1.25) %>% 
        height_all(height = 0.35) %>% 
        theme_zebra() %>% 
        align(align = "center", part = "all")
      
      # lms flex ----
      flextable_lms <- flextable(lms.pivots.df()) %>% 
        # colformat_num(col_keys = c("Location", "# Compliant Employees", "Total Employees", "Percent Compliant"), digits = 0) %>% 
        width(width = 1.25) %>% 
        height_all(height = 0.35) %>% 
        theme_zebra() %>% 
        align(align = "center", part = "all")
      
      example_pp <- read_pptx() %>% 
        add_slide(layout = "Title Slide", master = "Office Theme") %>% 
        ph_with_text(
          type = "ctrTitle",
          str = "This is a test"
        ) %>% 
        ph_with(
          location = ph_location_type(type = "subTitle"),
          value = "of the EHSS Weekly Automation Report download"
        ) %>% 
        
        # LMS slide ----
      add_slide(layout = "Title and Content", master = "Office Theme") %>% 
        ph_with_text(
          type = "title",
          str = paste("EHSS - Compliance Training Completion Status", months(Sys.Date() - months(1)), "/", months(Sys.Date()), year(Sys.Date()), "as of", format(Sys.Date(), format ="%m/%d/%Y"))
        ) %>% 
        ph_with_flextable_at(
          value = flextable_lms,
          left = 2.5,
          top = 2
        ) %>% 
        
        # Safety incidents slide ----
      add_slide(layout = "Title and Content", master = "Office Theme") %>% 
        ph_with_text(
          type = "title",
          str = "Safety - Incidents this week"
        ) %>% 
        ph_with_flextable_at(
          value = flextable_cbcs.inc,
          left = 2.5,
          top = 2
        ) %>% 
        
        # Claim cost/count slides (collapse to only one slide) ----
      add_slide(layout = "Title and Content", master = "Office Theme") %>% 
        ph_with_text(
          type = "title",
          str = paste(str_to_title(input$manager), "Territory - Claim Cost & Count -", Sys.Date())
        ) %>% 
        ph_with_flextable_at(
          value = flextable_cbcs.count,
          left = 2.5,
          top = 2
        ) %>% 
        
        add_slide(layout = "Title and Content", master = "Office Theme") %>% 
        ph_with_text(
          type = "title",
          str = paste(str_to_title(input$manager), "Territory - Claim Cost & Count -", Sys.Date())
        ) %>% 
        ph_with_flextable_at(
          value = flextable_cbcs.cost,
          left = 2.5,
          top = 2
        ) %>% 
        
        # Driver qual slide ----
      add_slide(layout = "Title and Content", master = "Office Theme") %>% 
        ph_with_text(
          type = "title",
          str = "Driver Qualification File Compliance Status",
        ) %>% 
        ph_with_flextable_at(
          value = flextable_lms,
          left = 2.5,
          top = 2
        )
      
      print(example_pp, target = file)
    }
  )
  
}

shinyApp(ui, server)
