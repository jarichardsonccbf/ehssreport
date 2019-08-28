library(readxl)
library(tidyverse)

source("data/locations.R")

CBCSList <- function(territory, week.start, week.end) {
  read_excel("data/CBCS_CCBF_LOSS_RUN.xls", skip = 6) %>% 
    rename(Dept.Name = `Dept Name`) %>% 
    left_join(cbcs.locations, by = "Dept.Name") %>% 
    mutate(Coverage = recode(Coverage,
                             "ALBI"  = "AUTO",
                             "ALPD"  = "AUTO",
                             "ALAPD" = "AUTO",
                             "GLPD"  = "GL"  ,
                             "GLBI"  = "GL")) %>% 
    filter(manager == territory) %>% 
    filter(`Occ Date` >= week.start & `Occ Date` <= week.end)
}

CBCSList(territory = "TAMPA", week.start = "2019-08-08", week.end = "2019-08-23")
