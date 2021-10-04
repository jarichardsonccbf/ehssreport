library(tidyverse)

a <- readRDS("lms_fac_name.Rds")

a <- a %>%
  mutate(Facility.Name = recode(Facility.Name,
                                "Lakeland Distribution Center" = "Winter Haven Distribution Center 3C03"),
         Facility.Name.Link = recode(Facility.Name.Link,
                                     "Lakeland Distribution Center - 3C03" = "Winter Haven Distribution Center - 3C03"))
write_rds(a, "lms_fac_name.Rds")
