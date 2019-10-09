library(tidyverse)
library(readxl)
library(lubridate)

source("locations2.R")

# stars.df <- read_excel("data/andrea10-4test/Stars.xlsx") %>% 
stars.df <- read_excel("data/STARSold.xlsx") %>% 
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
         Status != "Scheduled for Create",
         year == year(Sys.Date()),
         manager == "TAMPA") %>% 
  droplevels()

stars.df$Status


rm(stars.locations)

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
  
flextable_stars <- flextable(stars.pivot) %>% 
  add_header_lines(values = "STARS Status", top = TRUE) %>%
  border_remove() %>% 
  border(border.top = fp_border(color = "black"),
         border.bottom = fp_border(color = "black"),
         border.left = fp_border(color = "black"),
         border.right = fp_border(color = "black"), part = "all") %>% 
  align(part = "body", align = "center") %>% 
  align(part = "header", align = "center") %>% 
  align(j = 1, align = "left") %>% 
  align(part = "header", j = 1, align = "left") %>% 
  bold(bold = TRUE, part = "header") %>% 
  bold(bold = TRUE, part = "body", i = nrow(stars.pivot)) %>% 
  bold( i = ~ `Count of` == "Total") %>% 
  height(height = 0.23) %>% 
  width(width = 0.85, j = 2:6) %>% 
  width(width = 1.55, j = 1) %>% 
  height(height = 0.6, part = "header", i = 2) %>% 
  bg(bg = "light blue", part = "header") %>% 
  bg(bg = "light blue", part = "body", i = nrow(stars.pivot))
