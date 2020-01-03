library(tidyverse)
library(readxl)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

rm(cbcs.locations, jjkeller.locations, lms.locations)
 
stars.df <- read_excel("data/1-3-20STARS.xlsx") %>% 
  rename(Location = `Location Name`) %>% 
  left_join(stars.locations, by = "Location") %>%  
  mutate(year = year(`Creation Date`),
         `Investigation Type` = recode(`Investigation Type`,
                                       "Vehicle Incident Template" = "Vehicle",
                                       "Non-Vehicle Incident Template" = "Non-Vehicle"),
         `Status` = recode(`Status`,
                           "Complete - Nonpreventable" = "Complete NP",
                           "Complete - Preventable" = "Complete P",
                           "Complete - Non-preventable" = "Complete NP")) %>% 
  filter(Status != "Error Creating",
         `Investigation Type` == "Vehicle" | 
         `Investigation Type` == "Non-Vehicle",
         Status != "Scheduled for Create",
         year == year(Sys.Date())) %>% 
  droplevels()

rm(stars.locations)

one.loc <- stars.df %>% 
  filter(manager == "ORLANDO")
  
  
type <- one.loc %>% 
  group_by(Location, Status, `Investigation Type`) %>% 
  summarise (n = n()) %>% 
  pivot_wider(names_from = Status, values_from = n)

totals <- one.loc %>% 
  group_by(Location, Status) %>% 
  summarise(n = n()) %>% 
  pivot_wider(names_from = Status, values_from = n) %>% 
  mutate(`Investigation Type` = "A")

type.totals <- rbind(type, totals) %>%
  arrange(Location,
          `Investigation Type`) %>% 
  mutate(`Investigation Type` = recode(`Investigation Type`,
                                       "A" = "Total")) %>% 
  ungroup()

sums <- one.loc %>% 
  group_by(Status) %>% 
  summarise(n = n()) %>% 
  pivot_wider(names_from = Status, values_from = n) %>% 
  mutate(Total = "Total",
         empty = NA) %>% ungroup() %>% 
  select(Total, empty, everything())
 #  select(Total, empty, `Complete NP`, `Complete P`, New, `Pending IRC Review`)

colnames(sums) <- colnames(type.totals)

stars.pivot <- rbind(type.totals, sums) %>% 
  rename("Count of" = "Investigation Type")
  
