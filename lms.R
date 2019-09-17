library(tidyverse)

source("data/Locations.R")

lms <- read.csv("data/lms.csv")
  
lms %>% 
  left_join(lms.locations, "Org.Name") %>%
  filter(manager == "TAMPA") %>% 
  mutate(Item.Status = recode(Item.Status,
                              "In Progress" = "Incomplete",
                              "Not Started" = "Incomplete"),
         comp.binary = recode(Item.Status,
                              "Incomplete" = 0,
                              "Completed" = 1)) %>%
  group_by(Location, Title) %>%
  summarise(num.comp = sum(comp.binary),
            total = length(comp.binary),
            percent.compliant = num.comp/total * 100) %>% 
  select(Location, Title, percent.compliant) %>% 
  mutate(percent.compliant = round(percent.compliant, digits = 0),
         percent.compliant = paste(percent.compliant, "%", sep = "")) %>% 
  pivot_wider(names_from = Title, values_from = percent.compliant)
  

