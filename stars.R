library(tidyverse)
library(readxl)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

rm(cbcs.locations, jjkeller.locations, lms.locations, cbcs.cat, cbcs.pivot.cat, cbcs.years, typology.cost, typology.count)
 
stars.a <- read_excel("data/1-3-20STARS.xlsx") %>% 
  rename(Location = `Location Name`) %>% 
  left_join(stars.locations, by = "Location") %>% 
  filter(manager == "FT MYERS - BROWARD")

tie.in <- stars.locations %>% filter(manager != "PLACEHOLDER",
                                     manager == "FT MYERS - BROWARD") %>% select(Location) 
tie.in <- rbind(tie.in, tie.in, tie.in, tie.in, tie.in, tie.in, tie.in, tie.in) %>% arrange(Location)



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
  select(-c(Location, `Count of`)) %>% 
  mutate(Total = as.numeric(Total)) %>% 
  summarise_all(sum, na.rm = TRUE) %>% 
  mutate(Location = "Total",
         `Count of` = NA) %>% 
  select(Location, `Count of`, everything())

stars.pivot <- rbind(stars.pivot, col.tots)
