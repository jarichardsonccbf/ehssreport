library(readxl)
library(tidyverse)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

df <- read_excel("data/1-3-20CBCS_CCBF_LOSS_RUN 010320.xls", skip = 6) %>%
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
  filter(manager == "FT MYERS - BROWARD")


# claim cost 

claim.cost <- df %>% 
  group_by(`Loc Name`, Coverage) %>% 
  summarise(Incurred = sum(Incurred)) %>%
  ungroup() %>% 
  spread(Coverage, Incurred) %>%
  replace(., is.na(.), 0) %>% 
  mutate(Totals = rowSums(.[-1]))
  
claim.cost.tot <- claim.cost %>% 
  select(-c(`Loc Name`)) %>% 
  summarise_all(sum) %>% 
  mutate(Total = "Total") %>% 
  select(Total, everything())


colnames(claim.cost.tot) <- colnames(claim.cost)
  
# claim cost table
claim.cost <- rbind(claim.cost, claim.cost.tot)



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
