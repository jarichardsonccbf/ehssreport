library(tidyverse)
library(readxl)
library(lubridate)

source("locations2.R")

stars.df <- read_excel("data/STARS.xlsx") %>% 
  rename(Location = `Location Name`) %>% 
  left_join(stars.locations, by = "Location") %>%  
  mutate(year = year(`Creation Date`),
         `Investigation Type` = recode(`Investigation Type`,
                                       "Vehicle Incident Template" = "Vehicle",
                                       "Non-Vehicle Incident Template" = "Non-Vehicle"),
         `Status` = recode(`Status`,
                           "Complete - Nonpreventable" = "Complete NP",
                           "Complete - Preventable" = "Complete P")) %>% 
  filter(Status != "Error Creating",
         Status != "Scheduled for Create",
         year == year(Sys.Date())) %>% 
  droplevels()

type <- stars.df %>% 
  group_by(Location, Status, `Investigation Type`) %>% 
  summarise (n = n()) %>% 
  pivot_wider(names_from = Status, values_from = n)

totals <- stars.df %>% 
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

sums <- stars.df %>% 
  group_by(Status) %>% 
  summarise(n = n()) %>% 
  pivot_wider(names_from = Status, values_from = n) %>% 
  mutate(Total = "Total",
         empty = NA) %>% ungroup() %>% 
  select(Total, empty, `Complete NP`, `Complete P`, New, `Pending IRC Review`)

colnames(sums) <- colnames(type.totals)

stars.pivot <- rbind(type.totals, sums) %>% 
  rename("Count of" = "Investigation Type")
  mutate(Total = rowSums(.[3:6], na.rm = TRUE))
