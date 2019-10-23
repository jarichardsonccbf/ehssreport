library(readxl)
library(tidyverse)

source("locations2.R")

rm(jjkeller.locations, lms.locations, stars.locations)
  
cbcs.state <- read_excel("data/CBCS_CCBF_LOSS_RUN.xls", skip = 6) %>% 
  rename(Dept.Name = `Dept Name`) %>%  # Only if using Xls
  left_join(cbcs.locations, by = "Dept.Name")
  
cbcs.state <- cbcs.state %>% mutate(Coverage = recode(Coverage,
                                    "ALBI"  = "AUTO",
                                    "ALPD"  = "AUTO",
                                    "ALAPD" = "AUTO",
                                    "GLPD"  = "GL"  ,
                                    "GLBI"  = "GL") ,
                                    Incident = paste(`Occ Date`, "-",Coverage, "-",`Acc Desc`)) %>% 
  filter(as.Date(`Occ Date`, format = "%m/%d/%Y") >= "2019-08-08" & as.Date(`Occ Date`, format = "%m/%d/%Y") <= "2019-08-23",
         manager != "PLACEHOLDER") %>% 
  select(manager, Incident) %>% 
  group_by(manager) %>% 
  group_split()

cbcs.state[1]
