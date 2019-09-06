library(readxl)
library(tidyverse)
library(lubridate)

source("data/locations.R")


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


CBCSPivots("TAMPA")
