library(tidyverse)
library(flextable)
library(officer)
library(lubridate)
library(readxl)

source("locations2.R")

lms <- read_excel("data/Training Report 01162020.xlsx", sheet = "Full Data")

lms.pivots.df <- lms %>% 
  left_join(lms.locations, "Location") %>% 
  mutate(`Course Enrollment Status` = recode(`Course Enrollment Status`,
                              "Enrolled" = "Incomplete",
                              "In Progress" = "Incomplete"),
         comp.binary = recode(`Course Enrollment Status`,
                              "Incomplete" = 0,
                              "Completed" = 1)) %>%
  group_by(manager, location, `Employee Course`) %>%
  summarise(num.comp = sum(comp.binary),
            total = length(comp.binary),
            percent.compliant = num.comp/total * 100) %>% 
  select(location, `Employee Course`, percent.compliant) %>% 
  mutate(percent.compliant = round(percent.compliant, digits = 0),
         percent.compliant = paste(percent.compliant, "%", sep = "")) %>% 
  pivot_wider(names_from = `Employee Course`, values_from = percent.compliant) %>% 
  group_by(manager)



