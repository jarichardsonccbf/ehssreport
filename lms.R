library(tidyverse)
library(flextable)
library(officer)
library(lubridate)
library(readxl)

source("locations2.R")

lms <- read_excel("data/Weekly EHSS Monthly Completion_1-23-2020-121011-AM.xlsx", skip = 1)

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





lms <- read_excel("data/Weekly EHSS Monthly Completion_1-23-2020-121011-AM.xlsx", skip = 1) %>%
  left_join(lms.locations, "Location") %>%
  filter(manager != "PLACEHOLDER",
         manager == "FT MYERS - BROWARD")

lms %>% group_by(location, `Employee Course`, `Course Enrollment Status`) %>%
  summarise(n = length(`Course Enrollment Status`)) %>%
  ungroup %>% group_by(location, `Employee Course`) %>%
  mutate(percent.compliant = n / sum(n) * 100) %>%
  mutate(percent.compliant = round(percent.compliant, digits = 0),
         percent.compliant = paste(percent.compliant, "%", sep = "")) %>%
  select(-c(n)) %>%
  rename(`Status` = `Course Enrollment Status`,
         Location = location) %>%
  pivot_wider(names_from = `Employee Course`, values_from = percent.compliant)
