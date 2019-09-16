library(tidyverse)

source("data/locations.R")

lms <- read.csv("data/lms.csv")
  
a <- lms %>% 
  left_join(lms.locations, "Org.Name") %>%
  filter(manager == "TAMPA") %>% 
  mutate(Item.Status = recode(Item.Status,
                              "In Progress" = "Incomplete",
                              "Not Started" = "Incomplete"),
         comp.binary = recode(Item.Status,
                              "Incomplete" = 0,
                              "Completed" = 1)) %>%
  group_by(location, Title) %>%
  summarise(num.comp = sum(comp.binary),
            total = length(comp.binary),
            percent.compliant = num.comp/total * 100) %>% 
  select(location, Title, percent.compliant) %>% 
  pivot_wider(names_from = Title, values_from = percent.compliant)

