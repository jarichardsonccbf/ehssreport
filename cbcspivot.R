library(readxl)
library(tidyverse)
library(lubridate)

source("locations2.R")


CBCSPivots <- function(territory) {
  df <- read_excel("data/CBCS_CCBF_LOSS_RUN.xls", skip = 6) %>%
    mutate(Coverage = recode(Coverage,
                             "ALBI"  = "AUTO",
                             "ALPD"  = "AUTO",
                             "ALAPD" = "AUTO",
                             "GLPD"  = "GL"  ,
                             "GLBI"  = "GL"),
           occ.year = year(`Occ Date`)) %>%
    rename(Dept.Name = `Dept Name`) %>% 
    left_join(cbcs.locations, by = "Dept.Name") %>% 
    filter(manager == territory,
           occ.year == format(Sys.Date(), "%Y")) %>% 
    select(`Loc Name`, Coverage, Incurred)

  
# claim cost 
  
claim.cost <- df %>% 
  group_by(`Loc Name`, Coverage) %>% 
  summarise(Incurred = sum(Incurred)) %>%
  ungroup() %>% 
  spread(Coverage, Incurred) %>%
  replace(., is.na(.), 0) %>% 
  mutate(Totals = rowSums(.[2:4]))
  
claim.cost.tot <- data.frame("Total", sum(claim.cost$AUTO), sum(claim.cost$GL), sum(claim.cost$WC), sum(claim.cost$Totals))
  
colnames(claim.cost.tot) <- colnames(claim.cost)
  
# claim cost table
claim.cost <- rbind(claim.cost, claim.cost.tot)
  

# claim count

claim.count <- df %>% 
  group_by(`Loc Name`, Coverage) %>% 
  summarise (n = n()) %>% 
  spread(Coverage, n) %>% 
  replace(., is.na(.), 0) %>% 
  mutate(Totals = sum(AUTO, GL, WC)) %>% 
  ungroup()
  
claim.count.tot <- data.frame("Total", sum(claim.count$AUTO), sum(claim.count$GL), sum(claim.count$WC), sum(claim.count$Totals))
  
colnames(claim.count.tot) <- colnames(claim.count)
  
# claim count table
claim.count <- rbind(claim.count,claim.count.tot)

return(list(claim.cost, claim.count))  
  
}

flextable_cost <- flextable(CBCSPivots("TAMPA")[[1]]) %>% 
  add_header_lines(values = "Claim Cost", top = TRUE) %>% 
  border_remove() %>% 
  border(border.top = fp_border(color = "black"),
         border.bottom = fp_border(color = "black"),
         border.left = fp_border(color = "black"),
         border.right = fp_border(color = "black"), part = "all") %>%
  bold(bold = TRUE, part = "header") %>% 
  align(align = "left", part = "all", j = 1) %>% 
  align(align = "center", part = "header", j = 2:5) %>% 
  bg(bg = "light blue", part = "header") %>% 
  bg(bg = "light blue", part = "body", i = nrow(CBCSPivots("TAMPA")[[1]])) %>% 
  bold(bold = TRUE, part = "body", i = nrow(CBCSPivots("TAMPA")[[1]])) %>% 
  width(width = 1.48, j = 1) %>% 
  width(width = 1, j = 2:4) %>% 
  width(width = 1.14, j = 5)

flextable_count <- flextable(CBCSPivots("TAMPA")[[2]]) %>% 
  add_header_lines(values = "Claim Count", top = TRUE) %>% 
  border_remove() %>% 
  border(border.top = fp_border(color = "black"),
         border.bottom = fp_border(color = "black"),
         border.left = fp_border(color = "black"),
         border.right = fp_border(color = "black"), part = "all") %>%
  bold(bold = TRUE, part = "header") %>% 
  align(align = "left", part = "all", j = 1) %>% 
  align(align = "center", part = "header", j = 2:5) %>% 
  bg(bg = "light blue", part = "header") %>% 
  bg(bg = "light blue", part = "body", i = nrow(CBCSPivots("TAMPA")[[2]])) %>% 
  bold(bold = TRUE, part = "body", i = nrow(CBCSPivots("TAMPA")[[2]])) %>% 
  width(width = 1.48, j = 1) %>% 
  width(width = 1, j = 2:4) %>% 
  width(width = 1.14, j = 5)

cbcs.pivots.title <- "Safety"

cbcs.pivots.text <- paste("Tampa Territory - Claim Cost & Count -", format(Sys.Date(), format ="%m/%d/%Y"))
