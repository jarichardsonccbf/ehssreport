library(readxl)
library(tidyverse)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

rm(stars.locations, jjkeller.locations, lms.locations, cbcs.cat, cbcs.years, typology.cost, typology.count)

# df <- read_excel("data/1-3-20CBCS_CCBF_LOSS_RUN 010320.xls", skip = 6) %>%
  # mutate(Coverage = recode(Coverage,
  #                            "ALBI"  = "AUTO",
  #                            "ALPD"  = "AUTO",
  #                            "ALAPD" = "AUTO",
  #                            "GLPD"  = "GL"  ,
  #                            "GLBI"  = "GL"),
          # occ.year = year(`Occ Date`)) %>%
 # rename(Dept.Name = `Dept Name`) %>%
 # left_join(cbcs.locations, by = "Dept.Name") %>%
 # filter(manager != "PLACEHOLDER")


# cbcs.locations <- df %>%
#   select(`Loc Name`, manager) %>%
#   unique()

# coverage <- data.frame(Coverage = c("AUTO", "GL", "WC"), loc = rbind(cbcs.locations, cbcs.locations, cbcs.locations)) %>%
#   rename(manager = loc.manager,
#          Loc.Name = loc.Loc.Name) %>%
#   arrange(manager, Loc.Name, Coverage)

# cbcs.pivots <- df %>%
#   filter(occ.year == format(Sys.Date(), "%Y")) %>%
#   select(manager, `Loc Name`, Coverage, Incurred) %>%
#   rename(Loc.Name = `Loc Name`)
# 
# cbcs.pivots <- coverage %>%
#   left_join(cbcs.pivots, c("manager", "Loc.Name", "Coverage")) %>%
#   filter(manager == "FT MYERS - BROWARD")
# claim cost 

claim.cost <- df %>% 
  group_by(`Loc Name`, Coverage) %>% 
  summarise(Incurred = sum(Incurred)) %>%
  ungroup() %>% 
  spread(Coverage, Incurred) %>%
  replace(., is.na(.), 0) %>% 
  mutate(Totals = rowSums(.[-1])) #%>% 
  rename(`Loc Name` = Loc.Name)
  
claim.cost.tot <- claim.cost %>% 
  select(-c(`Loc Name`)) %>% 
  summarise_all(sum) %>% 
  mutate(Total = "Total") %>% 
  select(Total, everything())


colnames(claim.cost.tot) <- colnames(claim.cost)
  
# claim cost table
claim.cost <- rbind(claim.cost, claim.cost.tot)


# claim count

claim.count <- df %>%
  group_by(Loc.Name, Coverage) %>% 
  summarise(Incurred = sum(!is.na(Incurred))) %>%
  spread(Coverage, Incurred) %>% 
  replace(., is.na(.), 0) %>% 
  ungroup() %>% 
  mutate(Totals = rowSums(.[-1])) %>% 
  rename(`Loc Name` = Loc.Name)
  
claim.count.tot <- claim.count %>% 
  select(-c(`Loc Name`)) %>% 
  summarise_all(sum) %>% 
  mutate(Totalr = "Total") %>% 
  select(Totalr, everything())

colnames(claim.count.tot) <- colnames(claim.count)
  
claim.count <- rbind(claim.count,claim.count.tot)
