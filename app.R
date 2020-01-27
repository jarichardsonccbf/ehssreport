library(shiny)
library(tidyverse)
library(janitor)
library(scales)
library(lubridate)
library(readxl)
library(flextable)
library(officer)
library(xlsx)

options(shiny.maxRequestSize=30*1024^2)

source("locations2.R")

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

  titlePanel(title=div(img(src="logo.jpg"), "EHSS Regional Manager Weekly Reporting Tool")),

  hr(),

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

      # training upload ----
      h4("Attach training xlsx"),
      fileInput2("file3", "File location", labelIcon = "folder-open-o",
                 accept = c(".csv"), progress = TRUE),

      # CBCS upload ----
      h4("Attach CBCS Loss Run xlsx"),
      fileInput2("file2", "File location", labelIcon = "folder-open-o",
                 accept = c(".xlsx"), progress = TRUE),

      # JJK upload ----
      h4("Attach JJ Keller csv"),
      fileInput2("file1", "File location", labelIcon = "folder-open-o",
                 accept = c("text/csv",
                            "text/comma-separated-values,text/plain",
                            ".csv"), progress = TRUE),

      # STARS upload ----
      h4("Attach STARS xlsx"),
      fileInput2("file4", "File location", labelIcon = "folder-open-o",
                 accept = c(".xlsx"), progress = TRUE),

      # All territories or not ----
      selectInput("allstate", label = h4("Entire territory?"),
                  choices = list("NO", "YES")),

      # Territory selection ----
      selectInput("manager", label = h4("Region"),
                  choices = list("TAMPA",
                                 "ORLANDO",
                                 "SOUTH FL",
                                 "JACKSONVILLE",
                                 "TAMPA",
                                 "FT MYERS - BROWARD")),

      # Date range for CBCS Desc ----
      dateRangeInput('dateRange',
                     label = h4("Input date range for weekly incident table"),
                     start = Sys.Date() - 6, end = Sys.Date()),

      # Download button ----
      downloadButton("download_powerpoint", "Download Tables to Powerpoint")

    ),

    # Main panel display. Use tabs, one for each slide ----
    mainPanel(

      # Define tabs
      tabsetPanel(type = "tabs",
                  # General intro tab
                  tabPanel("Intro",
                           h5(p("Download data from LMS, CBCS, JJ Keller, and STARS.",
                                span("Do not", style = "color:red"),
                                "change any information in your raw data pulls. Upload the indicated data sources to the left. You may need to select \"All file types\" from the dropdown menu next to the \"File name\" entry field to find your file.")),
                           h5(p("Indicate 'YES' under 'Entire territory?' if you would like to see the entire CCBF system or 'NO' to display a specific region.")),
                           h5(p("The date range will default to the past 7 days and will display CBCS results for this range unless otherwise specified.")),
                           h5(p("Please report any errors, changes in vendors or data sources, or location shifts to Jason Richardson (jarichardson@cocacolaflorida.com). Please attach any data files that are giving you issues.
                               "))
                  ),

                  # Training tab
                  tabPanel("Training", uiOutput(outputId = "lms.pivot")),

                  # Incident descriptions tab
                  tabPanel("Weekly Safety Incidents", uiOutput(outputId = "weekly.cbcs.incidents")),

                  # Costs and Counts tab
                  tabPanel("Claims Costs and Counts", uiOutput(outputId = "cost.pivot"),
                           uiOutput(outputId = "count.pivot"),
                           uiOutput(outputId = "ytd.cost"),
                           uiOutput(outputId = "ytd.count")),

                  # JJ Keller tab
                  tabPanel("JJ Keller",  uiOutput(outputId = "driver.qual.table"),
                           plotOutput(outputId = "jjk.compliance.pie"),
                           plotOutput(outputId = "jjk.driver.stats.pie")),

                  # STARS tab
                  tabPanel("STARS",  tableOutput(outputId = "stars.status"))
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

  # JJK driver stats pie chart ----

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

    jjk.qual.totals <- data.frame("Total:", sum(jjk.qual$num.compliant), sum(jjk.qual$total), round(sum(jjk.qual$num.compliant)/sum(jjk.qual$total) * 100,2))

    colnames(jjk.qual.totals) <- colnames(jjk.qual)

    # driver qual table
    rbind(jjk.qual,jjk.qual.totals) %>%
      mutate(
        num.compliant = format(round(num.compliant, 0), nsmall = 0),
        percent.compliant = paste(round(percent.compliant,2), "%", sep = ""),
        Assigned.Location = recode(Assigned.Location,
                                   "Tampa Equipment Services" = "Tampa ES",
                                   "Tampa Fleet Shop" = "Tampa Fleet",
                                   "Tampa Transportation" = "Tampa Transport")
      ) %>%
      rename(
        Location = Assigned.Location,
        `# Compliant Employees` = num.compliant,
        `Total Employees` = total,
        `Percent Compliant` = percent.compliant
      )
  })

  output$driver.qual.table <- renderUI({
    jjk.qual.table() %>%
      flextable() %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      align(align = "center", part = "all") %>%
      align(align = "left", part = "body", j = 1) %>%
      bold(bold = TRUE, part = "body", i = nrow(flextable(jjk.qual.table())$body$dataset)) %>%
      bold(bold = TRUE, part = "header") %>%
      height(height = 0.74, part = "header") %>%
      height(height = 0.28, part = "body") %>%
      width(width = 1.4, j = 1) %>%
      width(width = 1.2, j = 2:ncol(flextable(jjk.qual.table())$body$dataset)) %>%
      bg(bg = "dark red", part = "header") %>%
      color(color = "white", part = "header") %>%
      htmltools_value()
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
      rename(Location = `Loc Name`) %>%
      mutate(Coverage = recode(Coverage,
                               "ALBI"  = "Auto",
                               "ALPD"  = "Auto",
                               "ALAPD" = "A",
                               "GLPD"  = "GL"  ,
                               "GLBI"  = "GL") ,
             Incident = paste(format(`Occ Date`, format = "%m/%d/%y"), "-",Coverage, "-",`Acc Desc`)) %>%
      filter(`Occ Date` >= input$dateRange[1] & `Occ Date` <= input$dateRange[2]) %>%   # change to `Occ Date` if using Xls
      select(Location, Incident)
  })

  output$weekly.cbcs.incidents <- renderUI({
    weekly.incidents() %>%
      flextable() %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      align(align = "left", part = "all") %>%
      height(part = "header", height = 0.33) %>%
      height(part = "body", height = 1) %>%
      width(width = 1.81, j = 1) %>%
      width(width = 5.59, j = 2) %>%
      htmltools_value()
  })

  # DF for CBCS pivots ----

  cbcs.pivots.df <- reactive({
    req(input$file2)

    if(input$manager == "JACKSONVILLE") {

      df <- read_excel(input$file2$datapath, skip = 6) %>%
          rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        mutate(Coverage = recode(Coverage,
                                 "ALBI"  = "GL",
                                 "ALPD"  = "PD",
                                 "ALAPD" = "PD",
                                 "GLPD"  = "PD"  ,
                                 "GLBI"  = "GL"),
               `Loc Name` = recode(`Loc Name`,
                                   "TAMPA CABOT WAREHOUSE" = "TAMPA CABOT"),
               occ.year = year(as.Date(`Occ Date`, format = "%m/%d/%Y"))) %>%
          left_join(cbcs.locations, by = "Dept.Name") %>%
        filter(manager != "PLACEHOLDER")

      cbcs.locations <- df %>%
        select(`Loc Name`, manager) %>%
        unique()

      coverage <- data.frame(Coverage = c("GL", "PD", "WC"), loc = rbind(cbcs.locations, cbcs.locations, cbcs.locations)) %>%
        rename(manager = loc.manager,
               Loc.Name = loc.Loc.Name) %>%
        arrange(manager, Loc.Name, Coverage)

      cbcs.pivots <- df %>%
        filter(occ.year == format(Sys.Date(), "%Y")) %>%
        select(manager, `Loc Name`, Coverage, Incurred) %>%
        rename(Loc.Name = `Loc Name`)

      cbcs.pivots <- coverage %>%
        left_join(cbcs.pivots, c("manager", "Loc.Name", "Coverage")) %>%
        filter(manager == input$manager)

    } else {

    if(input$allstate == "YES") {

      df <- read_excel(input$file2$datapath, skip = 6) %>%
        rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        mutate(Coverage = recode(Coverage,
                                 "ALBI"  = "Auto",
                                 "ALPD"  = "Auto",
                                 "ALAPD" = "Auto",
                                 "GLPD"  = "GL"  ,
                                 "GLBI"  = "GL"),
               `Loc Name` = recode(`Loc Name`,
                                   "TAMPA CABOT WAREHOUSE" = "TAMPA CABOT"),
               occ.year = year(as.Date(`Occ Date`, format = "%m/%d/%Y"))) %>%
        left_join(cbcs.locations, by = "Dept.Name") %>%
        filter(manager != "PLACEHOLDER")

      cbcs.locations <- df %>%
        select(`Loc Name`, manager) %>%
        unique()

      coverage <- data.frame(Coverage = c("Auto", "GL", "WC"), loc = rbind(cbcs.locations, cbcs.locations, cbcs.locations)) %>%
        rename(manager = loc.manager,
               Loc.Name = loc.Loc.Name) %>%
        arrange(manager, Loc.Name, Coverage)

      cbcs.pivots <- df %>%
        filter(occ.year == format(Sys.Date(), "%Y")) %>%
        select(manager, `Loc Name`, Coverage, Incurred) %>%
        rename(Loc.Name = `Loc Name`)

      cbcs.pivots <- coverage %>%
        left_join(cbcs.pivots, c("manager", "Loc.Name", "Coverage"))

    } else {

      df <- read_excel(input$file2$datapath, skip = 6) %>%
        rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
        mutate(Coverage = recode(Coverage,
                                 "ALBI"  = "Auto",
                                 "ALPD"  = "Auto",
                                 "ALAPD" = "Auto",
                                 "GLPD"  = "GL"  ,
                                 "GLBI"  = "GL"),
               `Loc Name` = recode(`Loc Name`,
                                   "TAMPA CABOT WAREHOUSE" = "TAMPA CABOT"),
               occ.year = year(as.Date(`Occ Date`, format = "%m/%d/%Y"))) %>%
        left_join(cbcs.locations, by = "Dept.Name") %>%
        filter(manager != "PLACEHOLDER")

      cbcs.locations <- df %>%
        select(`Loc Name`, manager) %>%
        unique()

      coverage <- data.frame(Coverage = c("Auto", "GL", "WC"), loc = rbind(cbcs.locations, cbcs.locations, cbcs.locations)) %>%
        rename(manager = loc.manager,
               Loc.Name = loc.Loc.Name) %>%
        arrange(manager, Loc.Name, Coverage)

      cbcs.pivots <- df %>%
        filter(occ.year == format(Sys.Date(), "%Y")) %>%
        select(manager, `Loc Name`, Coverage, Incurred) %>%
        rename(Loc.Name = `Loc Name`)

      cbcs.pivots <- coverage %>%
        left_join(cbcs.pivots, c("manager", "Loc.Name", "Coverage")) %>%
        filter(manager == input$manager)

      }
    }
  })

  # Cost Pivot ----

  cost.pivot <- reactive({

    costs <- cbcs.pivots.df() %>%
        group_by(Loc.Name, Coverage) %>%
        summarise(Incurred = sum(Incurred)) %>%
        ungroup() %>%
        spread(Coverage, Incurred) %>%
        replace(., is.na(.), 0) %>%
        mutate(Totals = rowSums(.[-1])) %>%
      rename(`Loc Name` = Loc.Name)

      costs.tot <- costs %>%
        select(-c(`Loc Name`)) %>%
        summarise_all(sum) %>%
        mutate(Total = "Total") %>%
        select(Total, everything())

      colnames(costs.tot) <- colnames(costs)

      # claim cost table

      claim.cost <- rbind(costs, costs.tot)

      claim.cost.r <- apply(claim.cost[2:ncol(claim.cost)], 2, dollar)

      claim.cost <- cbind(claim.cost[1], as.data.frame(as.matrix(claim.cost.r))) %>%
        rename(Location = `Loc Name`)

  })

  output$cost.pivot <- renderUI({

    costs <- cost.pivot() %>%
      flextable() %>%
      add_header_lines(values = "Claim Cost", top = TRUE) %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      bold(bold = TRUE, part = "header") %>%
      align(align = "center", part = "all") %>%
      align(align = "left", part = "all", j = 1) %>%
      bg(bg = "light blue", part = "header") %>%
      bg(bg = "light blue", part = "body", i = nrow(flextable(cost.pivot())$body$dataset)) %>%
      bold(bold = TRUE, part = "body", i = nrow(flextable(cost.pivot())$body$dataset)) %>%
      width(width = 1.48, j = 1) %>%
      width(width = 1, j = 2:ncol(flextable(cost.pivot())$body$dataset)) %>%
      htmltools_value()
  })

  # Count Pivot----

  count.pivot <- reactive({

      counts <- cbcs.pivots.df() %>%
        group_by(Loc.Name, Coverage) %>%
        summarise(Incurred = sum(!is.na(Incurred))) %>%
        spread(Coverage, Incurred) %>%
        replace(., is.na(.), 0) %>%
        ungroup() %>%
        mutate(Totals = rowSums(.[-1])) %>%
        rename(`Loc Name` = Loc.Name)

      counts.tot <- counts %>%
        select(-c(`Loc Name`)) %>%
        summarise_all(sum) %>%
        mutate(Totalr = "Total") %>%
        select(Totalr, everything())

      colnames(counts.tot) <- colnames(counts)

     # claim count table
      claim.count <- rbind(counts,counts.tot)

      format_count <- function (x) {format(round(x, 0), nsmall = 0)}

      claim.count.r <- apply(claim.count[2:ncol(claim.count)], 2, format_count)

      claim.count <- cbind(claim.count[1], as.data.frame(as.matrix(claim.count.r))) %>%
        rename(Location = `Loc Name`)

  })

  output$count.pivot <- renderUI({
    counts <- count.pivot() %>%
      flextable() %>%
      add_header_lines(values = "Claim Count", top = TRUE) %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      bold(bold = TRUE, part = "header") %>%
      align(align = "center", part = "all") %>%
      align(align = "left", part = "all", j = 1) %>%
      bg(bg = "light blue", part = "header") %>%
      bg(bg = "light blue", part = "body", i = nrow(flextable(count.pivot())$body$dataset)) %>%
      bold(bold = TRUE, part = "body", i = nrow(flextable(count.pivot())$body$dataset)) %>%
      width(width = 1.48, j = 1) %>%
      width(width = 1, j = 2:ncol(flextable(count.pivot())$body$dataset)) %>%
      htmltools_value()
  })


  # DF for CBCS ltr pivots ----

  cbcs.ltr.df <- reactive({
    req(input$file2)

    if(input$manager == "JACKSONVILLE") {

      df <- read_excel(input$file2$datapath, skip = 6) %>%
        mutate(Coverage = recode(Coverage,
                                 "ALBI"  = "GL",
                                 "ALPD"  = "PD",
                                 "ALAPD" = "PD",
                                 "GLPD"  = "PD"  ,
                                 "GLBI"  = "GL"),
               occ.year = year(`Occ Date`)) %>%
        rename(Dept.Name = `Dept Name`) %>%
        left_join(cbcs.locations, by = "Dept.Name") %>%
        filter(manager != "PLACEHOLDER")

      cbcs.locations <- df %>%
        select(`Loc Name`, manager) %>%
        unique() %>%
        left_join(cbcs.years, "manager") %>%
        left_join(cbcs.cat.jax, "year") %>%
        rename(occ.year = year) %>%
        mutate(occ.year = as.character(occ.year),
               occ.year = as.numeric(occ.year))

      prior.year <- df %>%
        filter(occ.year == as.numeric(format(Sys.Date(), "%Y")) - 1,
               `Occ Date` <= input$dateRange[2] - 366) %>%
        select(occ.year, `Loc Name`, Coverage, Incurred)

      present.year <- df %>%
        filter(occ.year == format(Sys.Date(), "%Y"),
               `Occ Date` <= input$dateRange[1]) %>%
        select(occ.year, `Loc Name`, Coverage, Incurred)

      year <- cbcs.locations %>%
        left_join(prior.year, by = c("Loc Name", "Coverage", "occ.year"))

      year <- year %>%
        left_join(present.year, by = c("Loc Name", "Coverage", "occ.year")) %>%
        mutate(Incurred = pmax(Incurred.x, Incurred.y, na.rm = TRUE)) %>%
        select(-c(Incurred.x, Incurred.y))

      year.loc <- year %>%
        filter(manager == input$manager)

    } else {

      df <- read_excel(input$file2$datapath, skip = 6) %>%
        mutate(Coverage = recode(Coverage,
                                 "ALBI"  = "AUTO",
                                 "ALPD"  = "AUTO",
                                 "ALAPD" = "AUTO",
                                 "GLPD"  = "GL"  ,
                                 "GLBI"  = "GL"),
               occ.year = year(`Occ Date`)) %>%
        rename(Dept.Name = `Dept Name`) %>%
        left_join(cbcs.locations, by = "Dept.Name") %>%
        filter(manager != "PLACEHOLDER")

      cbcs.locations <- df %>%
        select(`Loc Name`, manager) %>%
        unique() %>%
        left_join(cbcs.years, "manager") %>%
        left_join(cbcs.cat, "year") %>%
        rename(occ.year = year) %>%
        mutate(occ.year = as.character(occ.year),
               occ.year = as.numeric(occ.year))

      prior.year <- df %>%
        filter(occ.year == as.numeric(format(Sys.Date(), "%Y")) - 1,
               `Occ Date` <= input$dateRange[2] - 366) %>%
        select(occ.year, `Loc Name`, Coverage, Incurred)

      present.year <- df %>%
        filter(occ.year == format(Sys.Date(), "%Y"),
               `Occ Date` <= input$dateRange[1]) %>%
        select(occ.year, `Loc Name`, Coverage, Incurred)

      year <- cbcs.locations %>%
        left_join(prior.year, by = c("Loc Name", "Coverage", "occ.year"))

      year <- year %>%
        left_join(present.year, by = c("Loc Name", "Coverage", "occ.year")) %>%
        mutate(Incurred = pmax(Incurred.x, Incurred.y, na.rm = TRUE)) %>%
        select(-c(Incurred.x, Incurred.y))

      if(input$allstate == "YES")

        year.loc <- year

      else

        year.loc <- year %>%
        filter(manager == input$manager)

    }

  })


  # cbcs ltr cost ----

  ltr.cost <- reactive({

    if(input$manager == "JACKSONVILLE") {

    ytd.location.totals.cost <- cbcs.ltr.df() %>%
      group_by(occ.year, `Loc Name`) %>%
      summarise(Incurred = sum(Incurred, na.rm = TRUE)) %>%
      mutate(Coverage = paste(occ.year, "Total")) %>%
      select(occ.year, `Loc Name`, Coverage, Incurred) %>%
      ungroup()

    year.loc.cost <- cbcs.ltr.df() %>%
      group_by(occ.year, `Loc Name`, Coverage) %>%
      summarise(Incurred = sum(Incurred)) %>%
      ungroup()

    ytd.comp.cost <- rbind(as.data.frame(year.loc.cost), as.data.frame(ytd.location.totals.cost)) %>%
      ungroup() %>%
      pivot_wider(names_from = c(Coverage, occ.year), values_from = Incurred) %>%
      replace(., is.na(.), 0) %>%
      mutate(`Fav/Unfav` = .[[8]] - .[[9]]) %>%
      `colnames<-`(c('Location',
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "PD", sep = " "),
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "PD", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
                     'Fav/Unfav')) %>%
      select(Location,
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "PD", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "PD", sep = " "),
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
             `Fav/Unfav`)

    cost.totals <- ytd.comp.cost %>%
      select(-c(Location)) %>%
      summarise_all(sum, na.rm = TRUE) %>%
      mutate(Location = "Total") %>%
      select(Location, everything())

    ytd.comp.cost <- rbind(ytd.comp.cost, cost.totals)

    ytd.comp.cost.r <- apply(ytd.comp.cost[2:ncol(ytd.comp.cost)], 2, dollar)

    ytd.comp.cost <- cbind(ytd.comp.cost[1], as.data.frame(as.matrix(ytd.comp.cost.r)))

    } else {

      ytd.location.totals.cost <- cbcs.ltr.df() %>%
        group_by(occ.year, `Loc Name`) %>%
        summarise(Incurred = sum(Incurred, na.rm = TRUE)) %>%
        mutate(Coverage = paste(occ.year, "Total")) %>%
        select(occ.year, `Loc Name`, Coverage, Incurred) %>%
        ungroup()

      year.loc.cost <- cbcs.ltr.df() %>%
        group_by(occ.year, `Loc Name`, Coverage) %>%
        summarise(Incurred = sum(Incurred)) %>%
        ungroup()

      ytd.comp.cost <- rbind(as.data.frame(year.loc.cost), as.data.frame(ytd.location.totals.cost)) %>%
        ungroup() %>%
        pivot_wider(names_from = c(Coverage, occ.year), values_from = Incurred) %>%
        replace(., is.na(.), 0) %>%
        mutate(`Fav/Unfav` = .[[8]] - .[[9]]) %>%
        `colnames<-`(c('Location',
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "),
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
                       'Fav/Unfav')) %>%
        select(Location,
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
               `Fav/Unfav`)

      cost.totals <- ytd.comp.cost %>%
        select(-c(Location)) %>%
        summarise_all(sum, na.rm = TRUE) %>%
        mutate(Location = "Total") %>%
        select(Location, everything())

      ytd.comp.cost <- rbind(ytd.comp.cost, cost.totals)

      ytd.comp.cost.r <- apply(ytd.comp.cost[2:ncol(ytd.comp.cost)], 2, dollar)

      ytd.comp.cost <- cbind(ytd.comp.cost[1], as.data.frame(as.matrix(ytd.comp.cost.r)))

    }

  })

  # cbcs ltr count ----

  ltr.count <- reactive({

    if(input$manager == "JACKSONVILLE") {

    ytd.location.totals.count <- cbcs.ltr.df() %>%
      group_by(occ.year, `Loc Name`) %>%
      summarise(Incurred = sum(!is.na(Incurred))) %>%
      mutate(Coverage = paste(occ.year, "Total")) %>%
      select(occ.year, `Loc Name`, Coverage, Incurred) %>%
      ungroup()

    year.loc.count <- cbcs.ltr.df() %>%
      group_by(occ.year, `Loc Name`, Coverage) %>%
      summarise(Incurred = sum(!is.na(Incurred))) %>%
      ungroup()

    ytd.comp.count <- rbind(as.data.frame(year.loc.count), as.data.frame(ytd.location.totals.count)) %>%
      ungroup() %>%
      pivot_wider(names_from = c(Coverage, occ.year), values_from = Incurred) %>%
      replace(., is.na(.), 0) %>%
      mutate(`Fav/Unfav` = .[[8]] - .[[9]]) %>%
      `colnames<-`(c('Location',
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "PD", sep = " "),
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "PD", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
                     paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
                     paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
                     'Fav/Unfav')) %>%
      select(Location,
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "PD", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "PD", sep = " "),
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
             paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
             `Fav/Unfav`)

    count.totals <- ytd.comp.count %>%
      select(-c(Location)) %>%
      summarise_all(sum, na.rm = TRUE) %>%
      mutate(Location = "Total") %>%
      select(Location, everything())

    ytd.comp.count <- rbind(ytd.comp.count, count.totals)

    } else {

      ytd.location.totals.count <- cbcs.ltr.df() %>%
        group_by(occ.year, `Loc Name`) %>%
        summarise(Incurred = sum(!is.na(Incurred))) %>%
        mutate(Coverage = paste(occ.year, "Total")) %>%
        select(occ.year, `Loc Name`, Coverage, Incurred) %>%
        ungroup()

      year.loc.count <- cbcs.ltr.df() %>%
        group_by(occ.year, `Loc Name`, Coverage) %>%
        summarise(Incurred = sum(!is.na(Incurred))) %>%
        ungroup()

      ytd.comp.count <- rbind(as.data.frame(year.loc.count), as.data.frame(ytd.location.totals.count)) %>%
        ungroup() %>%
        pivot_wider(names_from = c(Coverage, occ.year), values_from = Incurred) %>%
        replace(., is.na(.), 0) %>%
        mutate(`Fav/Unfav` = .[[8]] - .[[9]]) %>%
        `colnames<-`(c('Location',
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "),
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
                       paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
                       paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
                       'Fav/Unfav')) %>%
        select(Location,
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
               paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
               paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
               `Fav/Unfav`)

      count.totals <- ytd.comp.count %>%
        select(-c(Location)) %>%
        summarise_all(sum, na.rm = TRUE) %>%
        mutate(Location = "Total") %>%
        select(Location, everything())

      ytd.comp.count <- rbind(ytd.comp.count, count.totals)

    }

  })

  # cbcs ltr outputs ----

  output$ytd.cost <- renderUI({

    ltr.cost() %>%
      mutate_if(is.numeric, round, 0) %>%
      flextable() %>%
      set_header_df(mapping = typology.cost, key = "col_keys") %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      align(align = "center", part = "all") %>%
      align(align = "left", part = "all", j = 1) %>%
      font(fontname = "arial", part = "all") %>%
      fontsize(size = 11, part = "all") %>%
      merge_at(i = 1, j = 2:3, part = "header") %>%
      merge_at(i = 1, j = 4:5, part = "header") %>%
      merge_at(i = 1, j = 6:7, part = "header") %>%
      merge_at(i = 1, j = 8:9, part = "header") %>%
      bold(bold = TRUE, part = "header") %>%
      bg(bg = "light blue", part = "header") %>%
      bg(bg = "light blue", part = "body", i = nrow(flextable(ltr.cost())$body$dataset)) %>%
      bold(bold = TRUE, part = "body", i = nrow(flextable(ltr.cost())$body$dataset)) %>%
      width(width = 1.48, j = 1) %>%
      width(width = 1, j = 2:ncol(flextable(ltr.cost())$body$dataset)) %>%
      htmltools_value()

  })

  output$ytd.count <- renderUI({

    ltr.count() %>%
      flextable() %>%
      set_header_df(mapping = typology.count, key = "col_keys") %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      align(align = "center", part = "all") %>%
      align(align = "left", part = "all", j = 1) %>%
      font(fontname = "arial", part = "all") %>%
      fontsize(size = 11, part = "all") %>%
      merge_at(i = 1, j = 2:3, part = "header") %>%
      merge_at(i = 1, j = 4:5, part = "header") %>%
      merge_at(i = 1, j = 6:7, part = "header") %>%
      merge_at(i = 1, j = 8:9, part = "header") %>%
      bold(bold = TRUE, part = "header") %>%
      bg(bg = "light blue", part = "header") %>%
      bg(bg = "light blue", part = "body", i = nrow(flextable(ltr.count())$body$dataset)) %>%
      bold(bold = TRUE, part = "body", i = nrow(flextable(ltr.count())$body$dataset)) %>%
      width(width = 1.48, j = 1) %>%
      width(width = 1, j = 2:ncol(flextable(ltr.count())$body$dataset)) %>%
      htmltools_value()

  })

  # LMS Pivot ----
  lms.pivots.df <- reactive({

    req(input$file3)

    if(input$manager == "FT MYERS - BROWARD"){

      if(input$allstate == "YES")

        lms <- read_excel(input$file3$datapath, skip = 1) %>%
          left_join(lms.locations, "Location") %>%
          filter(manager != "PLACEHOLDER")

      else

        lms <- read_excel(input$file3$datapath, skip = 1) %>%
          left_join(lms.locations, "Location") %>%
          filter(manager != "PLACEHOLDER",
                 manager == "FT MYERS - BROWARD")

      lms %>% group_by(location, `Employee Course`, `Course Enrollment Status`) %>%
        summarise(n = length(`Course Enrollment Status`)) %>%
        ungroup %>% group_by(location, `Employee Course`) %>%
        mutate(percent.compliant = n / sum(n) * 100) %>%
        mutate(percent.compliant = round(percent.compliant, digits = 0),
               percent.compliant = paste(percent.compliant, "%", sep = "")) %>%
        select(-c(n)) %>%
        rename(`Status` = `Course Enrollment Status`,
               Location = location) %>%
        pivot_wider(names_from = `Employee Course`, values_from = percent.compliant)

    } else {

      if(input$manager == "TAMPA"){

        if(input$allstate == "YES")

          lms <- read_excel(input$file3$datapath, skip = 1) %>%
            left_join(lms.locations, "Location") %>%
            filter(manager != "PLACEHOLDER")

        else

          lms <- read_excel(input$file3$datapath, skip = 1) %>%
            left_join(lms.locations, "Location") %>%
            filter(manager != "PLACEHOLDER",
                   manager == input$manager)

        lms %>%
          mutate(`Course Enrollment Status` = recode(`Course Enrollment Status`,
                                                            "Enrolled" = "Incomplete",
                                                            "In Progress" = "Incomplete"),
                       comp.binary = recode(`Course Enrollment Status`,
                                             "Incomplete" = 0,
                                             "Completed" = 1)) %>%
          group_by(manager, location, `Employee Course`) %>%
          summarise(num.comp = sum(comp.binary),
                    total = length(comp.binary),
                    percent.compliant = num.comp/total * 100) %>%
          select(location, `Employee Course`, percent.compliant) %>%
          mutate(percent.compliant = round(percent.compliant, digits = 0),
                 percent.compliant = paste(percent.compliant, "%", sep = "")) %>%
          pivot_wider(names_from = `Employee Course`, values_from = percent.compliant) %>%
          ungroup() %>%
          select(-c(manager)) %>%
          rename(Location = location)

      } else {

        if(input$allstate == "YES")

          lms <- read_excel(input$file3$datapath, skip = 1) %>%
            left_join(lms.locations, "Location") %>%
            filter(manager != "PLACEHOLDER")

        else

          lms <- read_excel(input$file3$datapath, skip = 1) %>%
            left_join(lms.locations, "Location") %>%
            filter(manager != "PLACEHOLDER",
                   manager == input$manager)

        lms %>%
          mutate(`Course Enrollment Status` = recode(`Course Enrollment Status`,
                                                           "Enrolled" = "Incomplete",
                                                           "In Progress" = "Incomplete"),
                       comp.binary = recode(`Course Enrollment Status`,
                                            "Incomplete" = 0,
                                            "Completed" = 1)) %>%
          group_by(manager, location, `Employee Course`) %>%
          summarise(num.comp = sum(comp.binary),
                    total = length(comp.binary),
                    percent.compliant = num.comp/total * 100) %>%
          select(location, `Employee Course`, percent.compliant) %>%
          mutate(percent.compliant = round(percent.compliant, digits = 0),
                 percent.compliant = paste(percent.compliant, "%", sep = "")) %>%
          pivot_wider(names_from = `Employee Course`, values_from = percent.compliant) %>%
          ungroup() %>%
          select(-c(manager)) %>%
          rename(Location = location)

      }

    }


  })

  output$lms.pivot <- renderTable({
    lms.pivots.df()
  })

  # STARS Pivot ----

  stars.pivots.df <- reactive({

    req(input$file4)

    if(input$allstate == "YES") {

      stars.a <- read_excel(input$file4$datapath) %>%
        rename(Location = `Location Name`) %>%
        left_join(stars.locations, by = "Location")

      tie.in <- stars.locations %>% filter(manager != "PLACEHOLDER") %>% select(Location)
      tie.in <- rbind(tie.in, tie.in, tie.in, tie.in, tie.in, tie.in, tie.in, tie.in) %>% arrange(Location)

    } else {

      stars.a <- read_excel(input$file4$datapath) %>%
        rename(Location = `Location Name`) %>%
        left_join(stars.locations, by = "Location") %>%
        filter(manager == input$manager)

      tie.in <- stars.locations %>% filter(manager != "PLACEHOLDER",
                                           manager == input$manager) %>% select(Location)
      tie.in <- rbind(tie.in, tie.in, tie.in, tie.in, tie.in, tie.in, tie.in, tie.in) %>% arrange(Location)

    }

    stars.b <- stars.a %>%
      mutate(year = year(`Creation Date`),
             `Investigation Type` = recode(`Investigation Type`,
                                           "Vehicle Incident Template" = "Vehicle",
                                           "Non-Vehicle Incident Template" = "Non-Vehicle"),
             `Status` = recode(`Status`,
                               "Complete - Nonpreventable" = "Complete NP",
                               "Complete - Preventable" = "Complete P",
                               "Complete - Non-preventable" = "Complete NP")) %>%
      filter(Status != "Error Creating",
             Status != "Scheduled for Create",
             `Investigation Type` == "Vehicle" |
               `Investigation Type` == "Non-Vehicle",
             year == year(Sys.Date())) %>%
      droplevels()

    type <- stars.b %>%
      group_by(Location, Status, `Investigation Type`) %>%
      summarise (n = n())

    tie.in <- data.frame(`Investigation Type` = c("Vehicle", "Vehicle", "Vehicle", "Vehicle", "Non-Vehicle", "Non-Vehicle", "Non-Vehicle", "Non-Vehicle"), Status = c("Complete NP", "Complete P", "New", "Pending IRC Review"), Location = tie.in) %>%
      arrange(Location, Status) %>%
      rename("Investigation Type" = Investigation.Type)

    type <- tie.in %>%
      left_join(type, by = c("Location", "Status", "Investigation Type")) %>%
      left_join(stars.locations, by = "Location") %>%
      # pivot_wider(names_from = Status, values_from = n) %>%
      replace(., is.na(.), 0)

    totals <- type %>%
      group_by(manager, Location, Status) %>%
      summarise(n = sum(n)) %>%
      # pivot_wider(names_from = Status, values_from = n) %>%
      mutate(`Investigation Type` = "A") %>%
      select(`Investigation Type`, Status, Location, n, manager)

    type.totals <- rbind(as.data.frame(type), as.data.frame(totals)) %>%
      arrange(Location,
              `Investigation Type`) %>%
      mutate(`Investigation Type` = recode(`Investigation Type`,
                                           "A" = "Total")) %>%
      ungroup() %>%
      select(Location, everything()) %>%
      pivot_wider(names_from = Status, values_from = n)

    stars.pivot <- type.totals %>%
      rename("Count of" = "Investigation Type") %>%
      mutate(Total = format(round(rowSums(.[4:ncol(type.totals)], na.rm = TRUE), 0), nsmall = 0)) %>%
      select(-c(manager))

    col.tots <- stars.pivot %>%
      filter(`Count of` == "Total") %>%
      select(-c(Location, `Count of`)) %>%
      mutate(Total = as.numeric(Total)) %>%
      summarise_all(sum, na.rm = TRUE) %>%
      mutate(Location = "Total",
             `Count of` = NA) %>%
      select(Location, `Count of`, everything())

    stars.pivot <- rbind(stars.pivot, col.tots)

  })

  output$stars.status <- renderUI({
    stars.pivots.df() %>%
      flextable() %>%
      add_header_lines(values = "STARS Status", top = TRUE) %>%
      border_remove() %>%
      border(border.top = fp_border(color = "black"),
             border.bottom = fp_border(color = "black"),
             border.left = fp_border(color = "black"),
             border.right = fp_border(color = "black"), part = "all") %>%
      align(part = "body", align = "center") %>%
      align(part = "header", align = "center") %>%
      align(j = 1, align = "left") %>%
      align(part = "header", j = 1, align = "left") %>%
      bold(bold = TRUE, part = "header") %>%
      bold(bold = TRUE, part = "body", i = nrow(flextable(stars.pivots.df())$body$dataset)) %>%
      bold( i = ~ `Count of` == "Total") %>%
      height(height = 0.23) %>%
      width(width = 0.85, j = 2:ncol(flextable(stars.pivots.df())$body$dataset)) %>%
      width(width = 1.55, j = 1) %>%
      height(height = 0.6, part = "header", i = 2) %>%
      bg(bg = "light blue", part = "header") %>%
      bg(bg = "light blue", part = "body", i = nrow(flextable(stars.pivots.df())$body$dataset)) %>%
      htmltools_value()
  })

  # PPT ----

  output$download_powerpoint <- downloadHandler(
    filename = function() {
      paste("P3Weekly_Summary_Deck_", input$manager, "_", gsub("-", "", Sys.Date()), ".pptx", sep = "")
    },
    content = function(file) {

      # jjk flex ----
      if (is.null(input$file1)) {

        NULL

      } else {

        flextable_dq <- flextable(jjk.qual.table()) %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          align(align = "center", part = "all") %>%
          align(align = "left", part = "body", j = 1) %>%
          bold(bold = TRUE, part = "body", i = nrow(flextable(jjk.qual.table())$body$dataset)) %>%
          bold(bold = TRUE, part = "header") %>%
          height(height = 0.74, part = "header") %>%
          height(height = 0.28, part = "body") %>%
          width(width = 1.4, j = 1) %>%
          width(width = 1.2, j = 2:ncol(flextable(jjk.qual.table())$body$dataset)) %>%
          bg(bg = "dark red", part = "header") %>%
          color(color = "white", part = "header")
      }

      # cbcs incidents flex ----
      if (is.null(input$file2)) {

        NULL

      } else {

        flextable_cbcs.inc <- flextable(weekly.incidents()) %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          align(align = "left", part = "all") %>%
          height(part = "header", height = 0.33) %>%
          height(part = "body", height = 0.55) %>%
          width(width = 1.81, j = 1) %>%
          width(width = 5.59, j = 2)
      }

      # cbcs cost flex ----
      if (is.null(input$file2)) {

        NULL

      } else {

        flextable_cbcs.cost <- flextable(cost.pivot()) %>%
          add_header_lines(values = "Claim Cost", top = TRUE) %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          bold(bold = TRUE, part = "header") %>%
          align(align = "left", part = "all", j = 1) %>%
          align(align = "center", part = "all") %>%
          align(align = "left", part = "all", j = 1) %>%
          bg(bg = "light blue", part = "header") %>%
          bg(bg = "light blue", part = "body", i = nrow(flextable(cost.pivot())$body$dataset)) %>%
          bold(bold = TRUE, part = "body", i = nrow(flextable(cost.pivot())$body$dataset)) %>%
          width(width = 1.48, j = 1) %>%
          width(width = 1, j = 2:ncol(flextable(cost.pivot())$body$dataset))
      }

      # cbcs count flex ----
      if (is.null(input$file2)) {

        NULL

      } else {

        flextable_cbcs.count <- flextable(count.pivot()) %>%
          add_header_lines(values = "Claim Count", top = TRUE) %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          bold(bold = TRUE, part = "header") %>%
          align(align = "center", part = "all") %>%
          align(align = "left", part = "all", j = 1) %>%
          bg(bg = "light blue", part = "header") %>%
          bg(bg = "light blue", part = "body", i = nrow(flextable(count.pivot())$body$dataset)) %>%
          bold(bold = TRUE, part = "body", i = nrow(flextable(count.pivot())$body$dataset)) %>%
          width(width = 1.48, j = 1) %>%
          width(width = 1, j = 2:ncol(flextable(count.pivot())$body$dataset))
      }

      # cbcs cost ltr flex ----
      if (is.null(input$file2)) {

        NULL

      } else {

        flextable_ltr.cost <- ltr.cost() %>%
          mutate_if(is.numeric, round, 0)

        flextable_ltr.cost <- flextable(flextable_ltr.cost) %>%
          set_header_df(mapping = typology.cost, key = "col_keys") %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          align(align = "center", part = "all") %>%
          align(align = "left", part = "all", j = 1) %>%
          font(fontname = "arial", part = "all") %>%
          fontsize(size = 11, part = "all") %>%
          merge_at(i = 1, j = 2:3, part = "header") %>%
          merge_at(i = 1, j = 4:5, part = "header") %>%
          merge_at(i = 1, j = 6:7, part = "header") %>%
          merge_at(i = 1, j = 8:9, part = "header") %>%
          bold(bold = TRUE, part = "header") %>%
          bg(bg = "light blue", part = "header") %>%
          bg(bg = "light blue", part = "body", i = nrow(flextable(ltr.cost())$body$dataset)) %>%
          bold(bold = TRUE, part = "body", i = nrow(flextable(ltr.cost())$body$dataset)) %>%
          width(width = 1.48, j = 1) %>%
          width(width = 1, j = 2:ncol(flextable(ltr.cost())$body$dataset))

      }

      # cbcs count ltr flex ----
      if (is.null(input$file2)) {

        NULL

      } else {

        flextable_ltr.count <- ltr.count() %>%
          flextable() %>%
          set_header_df(mapping = typology.count, key = "col_keys") %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          align(align = "center", part = "all") %>%
          align(align = "left", part = "all", j = 1) %>%
          font(fontname = "arial", part = "all") %>%
          fontsize(size = 11, part = "all") %>%
          merge_at(i = 1, j = 2:3, part = "header") %>%
          merge_at(i = 1, j = 4:5, part = "header") %>%
          merge_at(i = 1, j = 6:7, part = "header") %>%
          merge_at(i = 1, j = 8:9, part = "header") %>%
          bold(bold = TRUE, part = "header") %>%
          bg(bg = "light blue", part = "header") %>%
          bg(bg = "light blue", part = "body", i = nrow(flextable(ltr.count())$body$dataset)) %>%
          bold(bold = TRUE, part = "body", i = nrow(flextable(ltr.count())$body$dataset)) %>%
          width(width = 1.48, j = 1) %>%
          width(width = 1, j = 2:ncol(flextable(ltr.count())$body$dataset))

      }

      # lms flex ----
      if (is.null(input$file3)) {

        NULL

      } else {

        flextable_lms <- flextable(lms.pivots.df()) %>%
          border_remove() %>%
          rotate(rotation = "btlr", align = "center", part = "header", j = 2:length(flextable(lms.pivots.df())$col_keys)) %>%
          align(j = 1, part = "header") %>%
          align(j = 1, align = "left") %>%
          align(align = "center", part = "header") %>%
          add_header_lines(values = paste("EHSS Compliance Training ", months(Sys.Date() - months(1)), "-", months(Sys.Date())), top = TRUE) %>%
          height(part = "header", height = 2.28) %>%
          width(width = 1.7, j = 1) %>%
          width(width = 0.71, j = 2:length(flextable(lms.pivots.df())$col_keys)) %>%
          bg(bg = "light blue", part = "header") %>%
          height(height = 0.3, part = "header", i = 1) %>%
          bold(bold = TRUE, part = "header") %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all")
      }

      # stars flex ----
      if (is.null(input$file4)) {

        NULL

      } else {

        flextable_stars <- flextable(stars.pivots.df()) %>%
          add_header_lines(values = "STARS Status", top = TRUE) %>%
          border_remove() %>%
          border(border.top = fp_border(color = "black"),
                 border.bottom = fp_border(color = "black"),
                 border.left = fp_border(color = "black"),
                 border.right = fp_border(color = "black"), part = "all") %>%
          align(part = "body", align = "center") %>%
          align(part = "header", align = "center") %>%
          align(j = 1, align = "left") %>%
          align(part = "header", j = 1, align = "left") %>%
          bold(bold = TRUE, part = "header") %>%
          bold(bold = TRUE, part = "body", i = nrow(flextable(stars.pivots.df())$body$dataset)) %>%
          bold( i = ~ `Count of` == "Total") %>%
          height(height = 0.23) %>%
          width(width = 0.85, j = 2:ncol(flextable(stars.pivots.df())$body$dataset)) %>%
          width(width = 1.55, j = 1) %>%
          height(height = 0.6, part = "header", i = 2) %>%
          bg(bg = "light blue", part = "header") %>%
          bg(bg = "light blue", part = "body", i = nrow(flextable(stars.pivots.df())$body$dataset))
      }

      example_pp <- read_pptx() %>%
        add_slide(layout = "Title Slide", master = "Office Theme") %>%
        ph_with_text(
          type = "ctrTitle",
          str = "Weekly P3 Deck"
        ) %>%
        ph_with(
          location = ph_location_type(type = "subTitle"),
          value = "Copy and paste the generated tables into your report"
        )

      # LMS slide ----
      if (exists("flextable_lms")) {

        example_pp <- example_pp %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with(
            block_list(
              fpar(fp_p = fp_par(text.align = "center"),
                   ftext(paste("EHSS - Compliance Training Completion Status", " ", months(Sys.Date() - months(1)), "/", months(Sys.Date()), " ", year(Sys.Date()), " ","as of", " ",format(Sys.Date(), format ="%m/%d/%Y"), sep = ""),
                         prop = fp_text(font.size = 28)
                   )
              )
            ),
            location = ph_location_type(type = "title")) %>%
          ph_with_flextable(
            value = flextable_lms,
            type = "body"
          )
      }

      # Safety incidents slide ----
      if (exists("flextable_cbcs.inc")) {

        example_pp <- example_pp %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with_text(
            type = "title",
            str = "Safety - Incidents this week"
          ) %>%
          ph_with_flextable(
            value = flextable_cbcs.inc,
            type = "body"
          )
      }

      # Claim cost/count slides (collapse to only one slide) ----
      if (exists("flextable_cbcs.count") & exists("flextable_cbcs.cost")) {

        example_pp <- example_pp %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with(
            block_list(
              fpar(fp_p = fp_par(text.align = "center"),
                   ftext(paste(str_to_title(input$manager), "Territory - Claim Cost & Count -", format(Sys.Date(), format ="%m/%d/%Y")),
                         prop = fp_text(font.size = 20)
                   )
              )
            ),
            location = ph_location_type(type = "title")) %>%
          ph_with_flextable(
            value = flextable_cbcs.cost,
            type = "body"
          ) %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with(
            block_list(
              fpar(fp_p = fp_par(text.align = "center"),
                   ftext(paste(str_to_title(input$manager), "Territory - Claim Cost & Count -", format(Sys.Date(), format ="%m/%d/%Y")),
                         prop = fp_text(font.size = 20)
                   )
              )
            ),
            location = ph_location_type(type = "title")) %>%
          ph_with_flextable(
            value = flextable_cbcs.count,
            type = "body"
          )

      }

      # Claim cost/count ltr slides (collapse to only one slide) ----
      if (exists("flextable_ltr.cost") & exists("flextable_ltr.count")) {

        example_pp <- example_pp %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with(
            block_list(
              fpar(fp_p = fp_par(text.align = "center"),
                   ftext(paste("Claim Counts/Cost vs. Prior Year"),
                         prop = fp_text(font.size = 20)
                   )
              )
            ),
            location = ph_location_type(type = "title")) %>%
          ph_with_flextable(
            value = flextable_ltr.cost,
            type = "body"
          ) %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with(
            block_list(
              fpar(fp_p = fp_par(text.align = "center"),
                   ftext(paste("Claim Counts/Cost vs. Prior Year"),
                         prop = fp_text(font.size = 20)
                   )
              )
            ),
            location = ph_location_type(type = "title")) %>%
          ph_with_flextable(
            value = flextable_ltr.count,
            type = "body"
          )

      }

      # Driver qual slide ----
      if (exists("flextable_dq")) {

        example_pp <- example_pp %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with_text(
            type = "title",
            str = "Driver Qualification File Compliance Status",
          ) %>%
          ph_with_flextable(
            value = flextable_dq,
            type = "body"
          )
      }

      # stars slide ----
      if (exists("flextable_stars")) {

        example_pp <- example_pp %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
          ph_with(
            block_list(
              fpar(fp_p = fp_par(text.align = "center"),
                   ftext(paste("STARS Status - as of", format(Sys.Date(), format ="%m/%d/%Y")),
                         prop = fp_text(font.size = 20)
                   )
              )
            ),
            location = ph_location_type(type = "title")) %>%
          ph_with_flextable(
            value = flextable_stars,
            type = "body"
          )

      }

      print(example_pp, target = file)
    }
  )

}

shinyApp(ui, server)
