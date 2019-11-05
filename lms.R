library(tidyverse)
library(flextable)
library(officer)
library(lubridate)

source("locations2.R")

lms <- read.csv("data/jaxLMS.csv")

lms.pivots.df <- lms %>% 
  left_join(lms.locations, "Org.Name") %>% 
  mutate(Item.Status = recode(Item.Status,
                              "In Progress" = "Incomplete",
                              "Not Started" = "Incomplete"),
         comp.binary = recode(Item.Status,
                              "Incomplete" = 0,
                              "Completed" = 1)) %>%
  group_by(manager, location, Title) %>%
  summarise(num.comp = sum(comp.binary),
            total = length(comp.binary),
            percent.compliant = num.comp/total * 100) %>% 
  select(location, Title, percent.compliant) %>% 
  mutate(percent.compliant = round(percent.compliant, digits = 0),
         percent.compliant = paste(percent.compliant, "%", sep = "")) %>% 
  pivot_wider(names_from = Title, values_from = percent.compliant) %>% 
  group_by(manager) %>% 
  group_split()
  
lms.pivots.df[1]
a <- lms.pivots.df[2]
