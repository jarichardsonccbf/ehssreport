as.numeric(format(Sys.Date(), "%Y")) - 1
as.numeric(format(Sys.Date(), "%Y"))

library(readxl)
library(tidyverse)
library(lubridate)
library(flextable)
library(officer)

source("locations2.R")

rm(jjkeller.locations, lms.locations, stars.locations)

df <- read_excel("data/1-3-20CBCS_CCBF_LOSS_RUN 010320.xls", skip = 6) %>%
  mutate(Coverage = recode(Coverage,
                           "ALBI"  = "AUTO",
                           "ALPD"  = "AUTO",
                           "ALAPD" = "AUTO",
                           "GLPD"  = "GL"  ,
                           "GLBI"  = "GL"),
         occ.year = year(`Occ Date`)) %>%
  rename(Dept.Name = `Dept Name`) %>% 
  left_join(cbcs.locations, by = "Dept.Name") %>% 
  filter(manager != "PLACEHOLDER")

cbcs.locations <- df %>% 
  select(`Loc Name`, manager) %>% 
  unique() %>% 
  left_join(cbcs.years, "manager") %>% 
  left_join(cbcs.cat, "year") %>% 
  rename(occ.year = year) %>% 
  mutate(occ.year = as.character(occ.year),
         occ.year = as.numeric(occ.year))

prior.year <- df %>% 
  filter(occ.year == as.numeric(format(Sys.Date(), "%Y")) - 1,
         `Occ Date` <= Sys.Date() - 366) %>%
  select(occ.year, `Loc Name`, Coverage, Incurred)

present.year <- df %>% 
  filter(occ.year == format(Sys.Date(), "%Y"),
         `Occ Date` <= Sys.Date()) %>% 
  select(occ.year, `Loc Name`, Coverage, Incurred) 
  
rm(df, cbcs.cat, cbcs.years)

year <- cbcs.locations %>% 
  left_join(prior.year, by = c("Loc Name", "Coverage", "occ.year"))

year <- year %>% 
  left_join(present.year, by = c("Loc Name", "Coverage", "occ.year")) %>% 
  mutate(Incurred = pmax(Incurred.x, Incurred.y, na.rm = TRUE)) %>% 
  select(-c(Incurred.x, Incurred.y))

year.loc <- year %>% 
  filter(manager == "TAMPA")

# cost

ytd.location.totals.cost <- year.loc %>% 
  group_by(occ.year, `Loc Name`) %>% 
  summarise(Incurred = sum(Incurred, na.rm = TRUE)) %>% 
  mutate(Coverage = paste(occ.year, "Total")) %>% 
  select(occ.year, `Loc Name`, Coverage, Incurred) %>% 
  ungroup()

year.loc.cost <- year.loc %>% 
  group_by(occ.year, `Loc Name`, Coverage) %>% 
  summarise(Incurred = sum(Incurred)) %>%
  ungroup() 

ytd.comp.cost <- rbind(as.data.frame(year.loc.cost), as.data.frame(ytd.location.totals.cost)) %>%
  ungroup() %>% 
  pivot_wider(names_from = c(Coverage, occ.year), values_from = Incurred) %>% 
  replace(., is.na(.), 0) %>% 
  mutate(`Fav/Unfav` = .[[8]] - .[[9]]) %>% 
  `colnames<-`(c('Location', 
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "), 
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "), 
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
                 'Fav/Unfav')) %>% 
  select(Location,
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "), 
         paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "), 
         paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
         paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
         paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
         `Fav/Unfav`) 

cost.totals <- ytd.comp.cost %>% 
  select(-c(Location)) %>% 
  summarise_all(sum, na.rm = TRUE) %>% 
  mutate(Location = "Total") %>% 
  select(Location, everything())

ytd.comp.cost <- rbind(ytd.comp.cost, cost.totals)

format_money <- function (x) {paste("$", prettyNum(round(x, 0), big.mark = ","), sep = "")} 

ytd.comp.cost.r <- apply(ytd.comp.cost[2:ncol(ytd.comp.cost)], 2, format_money)

ytd.comp.cost <- cbind(ytd.comp.cost[1], as.data.frame(as.matrix(ytd.comp.cost.r)))


# count

ytd.location.totals.count <- year.loc %>% 
  group_by(occ.year, `Loc Name`) %>% 
  summarise(Incurred = sum(!is.na(Incurred))) %>% 
  mutate(Coverage = paste(occ.year, "Total")) %>% 
  select(occ.year, `Loc Name`, Coverage, Incurred) %>% 
  ungroup()

year.loc.count <- year.loc %>% 
  group_by(occ.year, `Loc Name`, Coverage) %>% 
  summarise(Incurred = sum(!is.na(Incurred))) %>% 
  ungroup() 

ytd.comp.count <- rbind(as.data.frame(year.loc.count), as.data.frame(ytd.location.totals.count)) %>%
  ungroup() %>% 
  pivot_wider(names_from = c(Coverage, occ.year), values_from = Incurred) %>% 
  replace(., is.na(.), 0) %>% 
  mutate(`Fav/Unfav` = .[[8]] - .[[9]]) %>% 
  `colnames<-`(c('Location', 
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "), 
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "), 
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
                 paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
                 paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
                 'Fav/Unfav')) %>% 
  select(Location,
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "AUTO", sep = " "), 
         paste(as.numeric(format(Sys.Date(), "%Y")), "AUTO", sep = " "),
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "GL", sep = " "), 
         paste(as.numeric(format(Sys.Date(), "%Y")), "GL", sep = " "),
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "WC", sep = " "),
         paste(as.numeric(format(Sys.Date(), "%Y")), "WC",  sep = " "),
         paste(as.numeric(format(Sys.Date(), '%Y')) - 1, "Total", sep = " "),
         paste(as.numeric(format(Sys.Date(), "%Y")), "Total", sep = " "),
         `Fav/Unfav`)

count.totals <- ytd.comp.count %>% 
  select(-c(Location)) %>% 
  summarise_all(sum, na.rm = TRUE) %>% 
  mutate(Location = "Total") %>% 
  select(Location, everything())

ytd.comp.count <- rbind(ytd.comp.count, count.totals)

rm(list= ls()[!(ls() %in% c('ytd.comp.cost','ytd.comp.count','typology.cost','typology.count'))])

ytd.comp.cost %>% 
  mutate_if(is.numeric, round, 0) %>% 
  flextable() %>% 
  set_header_df(mapping = typology.cost, key = "col_keys") %>% 
  border_remove() %>% 
  border(border.top = fp_border(color = "black"),
         border.bottom = fp_border(color = "black"),
         border.left = fp_border(color = "black"),
         border.right = fp_border(color = "black"), part = "all") %>% 
  align(align = "center", part = "all") %>% 
  align(align = "left", part = "all", j = 1) %>% 
  font(fontname = "arial", part = "all") %>% 
  fontsize(size = 11, part = "all") %>% 
  merge_at(i = 1, j = 2:3, part = "header") %>% 
  merge_at(i = 1, j = 4:5, part = "header") %>%
  merge_at(i = 1, j = 6:7, part = "header") %>% 
  merge_at(i = 1, j = 8:9, part = "header") %>% 
  bold(bold = TRUE, part = "header") %>% 
  bg(bg = "light blue", part = "header") %>% 
  bg(bg = "light blue", part = "body", i = nrow(ytd.comp.cost)) %>% 
  bold(bold = TRUE, part = "body", i = nrow(ytd.comp.cost)) %>% 
  width(width = 1.48, j = 1) %>% 
  width(width = 1, j = 2:ncol(ytd.comp.cost))

ytd.comp.count %>% 
  flextable() %>% 
  set_header_df(mapping = typology.count, key = "col_keys") %>% 
  border_remove() %>% 
  border(border.top = fp_border(color = "black"),
         border.bottom = fp_border(color = "black"),
         border.left = fp_border(color = "black"),
         border.right = fp_border(color = "black"), part = "all") %>% 
  align(align = "center", part = "all") %>% 
  align(align = "left", part = "all", j = 1) %>% 
  font(fontname = "arial", part = "all") %>% 
  fontsize(size = 11, part = "all") %>% 
  merge_at(i = 1, j = 2:3, part = "header") %>% 
  merge_at(i = 1, j = 4:5, part = "header") %>%
  merge_at(i = 1, j = 6:7, part = "header") %>% 
  merge_at(i = 1, j = 8:9, part = "header") %>% 
  bold(bold = TRUE, part = "header") %>% 
  bg(bg = "light blue", part = "header") %>% 
  bg(bg = "light blue", part = "body", i = nrow(ytd.comp.count)) %>% 
  bold(bold = TRUE, part = "body", i = nrow(ytd.comp.count)) %>% 
  width(width = 1.48, j = 1) %>% 
  width(width = 1, j = 2:ncol(ytd.comp.cost))

flextable_cbcs.cost.ltr
  