library(officer)

source("lms.R")
rm(cbcs.locations, jjkeller.locations, lms, lms.locations, lms.pivots.df, stars.locations)

example_pp <- read_pptx("2019 Weekly Safety P3 Deck - 092719.pptx")

slide7.structure <- slide_summary(example_pp, 7)

slide2.content <- pptx_summary(example_pp) %>% 
  filter(slide_id == 2)

example_pp <- example_pp %>% ph_remove(id = 3, type = "body")


print(example_pp, target = "filehsstest2.pptx")
