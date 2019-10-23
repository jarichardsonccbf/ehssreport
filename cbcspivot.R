library(readxl)
library(tidyverse)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

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
  filter(occ.year == format(Sys.Date(), "%Y"),
         manager != "PLACEHOLDER") %>%
  select(manager, `Loc Name`, Coverage, Incurred) %>% 
  group_by(manager) %>% 
  group_split(manager)


# claim cost 

ClaimCost <- function (manager.area) {
claim.cost <- manager.area %>% 
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

return(claim.cost)
  
}

ClaimCost(df[[4]])


# claim count

ClaimCount <- function (manager.area) {
  
  claim.count <- manager.area %>% 
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
  
  return(claim.count) 

}

ClaimCount(df[[5]])
