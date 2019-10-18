library(officer)

source("lms.R")
source("cbcsdesc.R")
source("cbcspivot.R")
source("jjkeller.R")
source("stars.R")

rm(cbcs.locations, jjkeller.locations, lms, lms.locations, lms.pivots.df, stars.locations, CBCSList, stars.df, stars.pivot, sums, totals, type, type.totals, CBCSPivots, JjkGraphs)

jjk.title <- cbcs.pivots.title
jjk.test <- "Driver Qualification File Compliance Status"

stars.title <- jjk.title
stars.text <- paste("STARS Status - as of", format(Sys.Date(), format ="%m/%d/%Y"))

example_pp <- read_pptx("2019 Weekly Safety P3 Deck - 092719.pptx")

# 3 lms
# 4 cbcs table
# 6 cbcs pivots
# 8 jjk
# 9 stars

print(example_pp, target = "filehsstest2.pptx")