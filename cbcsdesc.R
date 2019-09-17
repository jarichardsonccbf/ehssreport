library(readxl)
library(tidyverse)

source("data/locations.R")

CBCSList <- function(territory, week.start, week.end) {
  
  cbcs.state <- read_excel("data/CBCS_CCBF_LOSS_RUN.xls", skip = 6) %>% 
  rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
  left_join(cbcs.locations, by = "Dept.Name") %>% 
  filter(manager == territory)
  
  cbcs.state %>% mutate(Coverage = recode(Coverage,
                             "ALBI"  = "AUTO",
                             "ALPD"  = "AUTO",
                             "ALAPD" = "AUTO",
                             "GLPD"  = "GL"  ,
                             "GLBI"  = "GL") ,
                        Incident = paste(`Occ Date`, "-",Coverage, "-",`Acc Desc`)) %>% 
    filter(as.Date(`Occ Date`, format = "%m/%d/%Y") >= week.start & as.Date(`Occ Date`, format = "%m/%d/%Y") <= week.end) %>% 
    select(Incident)
}

CBCSList(territory = "TAMPA", week.start = "2019-08-08", week.end = "2019-08-23")
