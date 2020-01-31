library(readxl)
library(tidyverse)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

rm(stars.locations, jjkeller.locations, lms.locations, cbcs.cat, cbcs.years, typology.cost, typology.count, cbcs.pivot.cat.jax, typology.cost.jax, cbcs.cat.jax)


      df <- read_excel("data/Copy of CBCS_CCBF_LOSS_RUN 013120.xls", skip = 6) %>%
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
        filter(manager != "PLACEHOLDER") %>%
        filter(`Occ Date` >= input$dateRange[1] & `Occ Date` <= input$dateRange[2])

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
        filter(manager == "TAMPA")

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

