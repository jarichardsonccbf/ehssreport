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

      # CBCS upload ----
      h4("Attach CBCS Loss Run xlsx"),
      fileInput2("file2", "File location", labelIcon = "folder-open-o",
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

    ),

    # Main panel display. Use tabs, one for each slide ----
    mainPanel(

      # Define tabs
      tabsetPanel(type = "tabs",


                  # Costs and Counts tab
                  tabPanel("Claims Costs and Counts", uiOutput(outputId = "cost.pivot"),
                           uiOutput(outputId = "count.pivot"),
                           uiOutput(outputId = "ytd.cost"),
                           uiOutput(outputId = "ytd.count"))

      )

    )
  )
)

server <- function(input, output, session) {


  # DF for CBCS pivots ----



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

      year.loc <- year %>%
        filter(manager == "JACKSONVILLE")

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

        year.loc <- year %>%
        filter(manager == input$manager)


    }
  })


  # cbcs ltr cost ----

  ltr.cost <- reactive({

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
             paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep
                   = " "),
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

  })

  # cbcs ltr count ----

  ltr.count <- reactive({

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



}

shinyApp(ui, server)
