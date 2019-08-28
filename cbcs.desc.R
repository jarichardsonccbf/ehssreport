library(readxl)
library(tidyverse)

source("data/locations.R")

cbcs <- function(territory, week.start, week.end) {
  read_excel("data/CBCS_CCBF_LOSS_RUN.xls", skip = 6) %>% 
    rename(Dept.Name = `Dept Name`) %>% 
    left_join(location.key, by = "Dept.Name") %>% 
    mutate(Coverage = recode(Coverage,
                             "ALBI" = "AUTO",
                             "ALPD" = "AUTO",
                             "GLPB" = "GL"  ,
                             "GLBI" = "GL")) %>% 
    filter(manager == territory) %>% 
    filter(`Occ Date` >= week.start & `Occ Date` <= week.end)
}

cbcs(territory = "TAMPA", week.start = "2019-08-08", week.end = "2019-08-23")
