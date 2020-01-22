library(tidyverse)
library(janitor)
library(scales)

source("locations2.R")

df <- read.csv("data/jjkeller.csv") %>%
   left_join(jjkeller.locations, "Assigned.Location") %>%
   filter(DQ.File == "In Compliance" | DQ.File == "Out of Compliance") %>%
   mutate(dq.binary = recode(DQ.File,
                             "In Compliance" = 1,
                             "Out of Compliance" = 0)) %>%
   group_by(manager, Assigned.Location) %>%
   summarise(num.compliant     = sum(dq.binary      ),
             total             = length(dq.binary   ),
             percent.compliant = num.compliant/total * 100) %>% 
   group_by(manager) %>% 
   group_split() 
   
JjkQual <- function (manager.loc) {

   jjk.qual.totals <- data.frame("Total:", sum(manager.loc$num.compliant), sum(manager.loc$total), round(sum(manager.loc$num.compliant)/sum(manager.loc$total) * 100,2))
   
   colnames(jjk.qual.totals) <- colnames(manager.loc %>% select(-c(manager)))
   
   dq.table <- rbind(manager.loc %>% select(-c(manager)), jjk.qual.totals)

   return(dq.table)
}

JjkQual(df[[1]])






# DQ pie chart

dq.pie <- df %>%
 filter(DQ.File == "In Compliance" | DQ.File == "Out of Compliance") %>%
 group_by(DQ.File) %>%
 summarise (n = n()) %>%
 mutate(DQ.File = recode(DQ.File,
                         "In Compliance" = "Drivers In Compliance",
                         "Out of Compliance" = "Drivers Out Of Compliance"),
        freq = round((n / sum(n)) * 100, 2),
        label = paste(DQ.File, "-", paste(freq, "%", sep = ""))) %>%
 select(-c(n, DQ.File)) %>%
 ggplot(aes(x = 1, y = freq, fill = label)) +
 coord_polar(theta = 'y') +
 geom_bar(stat = "identity", color = 'black') +
 scale_fill_manual(values = c("darkgreen", "red")) +
 theme_minimal()+
 theme(
   axis.title.x = element_blank(),
   axis.text = element_blank(),
   axis.title.y = element_blank(),
   panel.border = element_blank(),
   panel.grid = element_blank(),
   axis.ticks = element_blank(),
   plot.title = element_text(size=14, face="bold"),
   legend.title = element_blank(),
   axis.text.x = element_blank(),
   legend.background = element_rect(linetype = "solid"))



# Driver Stats pie chart

dq.stats <- df %>%
 group_by(Status) %>%
 summarise (n = n()) %>%
 mutate(freq = round((n / sum(n)) * 100, 2),
        Status = recode(Status,
                        "Not Driving-" = "Not Driving"),
        label = paste(Status, "-", paste(freq, "%", sep = ""))) %>%
 select(-c(n, Status)) %>%
 ggplot(aes(x = 1, y = freq, fill = label)) +
 coord_polar(theta='y') +
 geom_bar(stat = "identity", color = 'black') +
 scale_fill_manual(values = c("deepskyblue4",
                              "firebrick4",
                              "yellowgreen",
                              "darkslateblue",
                              "darkcyan",
                              "chocolate3",
                              "lightsteelblue2",
                              "lightpink4",
                              "darkolivegreen4")) +
 theme_minimal()+
 theme(
   axis.title.x = element_blank(),
   axis.text = element_blank(),
   axis.title.y = element_blank(),
   panel.border = element_blank(),
   panel.grid = element_blank(),
   axis.ticks = element_blank(),
   plot.title = element_text(size=14, face="bold"),
   legend.title = element_blank(),
   axis.text.x = element_blank(),
   legend.background = element_rect(linetype = "solid")) +
 guides(fill = guide_legend(override.aes = list(colour=NA)))




library(flextable)
library(officer)

flextable_jjk <- flextable(JjkQual(df[[5]])) %>% 
   border_remove() %>% 
   border(border.top = fp_border(color = "black"),
          border.bottom = fp_border(color = "black"),
          border.left = fp_border(color = "black"),
          border.right = fp_border(color = "black"), part = "all") %>% 
   align(align = "center", part = "all") %>% 
   align(align = "left", part = "body", j = 1) %>% 
   bold(bold = TRUE, part = "body", i = nrow(JjkQual(df[[5]]))) %>% 
   bold(bold = TRUE, part = "header") %>% 
   height(height = 0.74, part = "header") %>% 
   height(height = 0.28, part = "body") %>% 
   width(width = 1.4, j = 1) %>% 
   width(width = 1.2, j = 2:4) %>% 
   bg(bg = "dark red", part = "header") %>% 
   color(color = "white", part = "header")
