library(tidyverse)
library(flextable)
library(officer)
library(lubridate)

source("locations2.R")

lms <- read.csv("data/orlandolms42.csv")

lms.pivots.df <- lms %>% 
  left_join(lms.locations, "Org.Name") %>% 
  filter(manager == "ORLANDO") %>% 
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
  mutate(percent.compliant = round(percent.compliant, digits = 0),
         percent.compliant = paste(percent.compliant, "%", sep = "")) %>% 
  pivot_wider(names_from = Title, values_from = percent.compliant)
  
flextable_lms <- flextable(lms.pivots.df) %>% 
  border_remove() %>% 
  rotate(rotation = "btlr", align = "center", part = "header", j = 2:length(flextable(lms.pivots.df)$col_keys)) %>% 
  align(j = 1, part = "header") %>% 
  align(j = 1, align = "left") %>%
  align(align = "center", part = "header") %>% 
  add_header_lines(values = paste("EHSS Compliance Training ", months(Sys.Date() - months(1)), "-", months(Sys.Date())), top = TRUE) %>% 
  height(part = "header", height = 2.28) %>%  
  width(width = 1.35, j = 1) %>% 
  width(width = 0.71, j = 2:length(flextable(lms.pivots.df)$col_keys)) %>% 
  bg(bg = "light blue", part = "header") %>% 
  height(height = 0.3, part = "header", i = 1) %>% 
  bold(bold = TRUE, part = "header") %>% 
  border(border.top = fp_border(color = "black"),
         border.bottom = fp_border(color = "black"),
         border.left = fp_border(color = "black"),
         border.right = fp_border(color = "black"), part = "all")

flextable_lms

