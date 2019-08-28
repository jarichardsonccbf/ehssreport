library(tidyverse)
library(janitor)

source("data/locations.R")

jj.keller <- read.csv("data/jjkeller.csv")

jj.keller.loc <- jj.keller %>% 
  left_join(jjkeller.locations, "Assigned.Location") %>% 
  filter(manager == "TAMPA",
         DQ.File == "In Compliance" | DQ.File == "Out of Compliance")

a <- jj.keller.loc %>%
  mutate(dq.binary = recode(DQ.File,
                            "In Compliance" = 1,
                            "Out of Compliance" = 0)) %>% 
  group_by(Assigned.Location) %>% 
  summarise(num.compliant     = sum(dq.binary      ),
            total             = length(dq.binary   ),
            percent.compliant = num.compliant/total*100) 

b <- data.frame("Total:", sum(a$num.compliant), sum(a$total), round(sum(a$num.compliant)/sum(a$total) * 100,2))

colnames(b) <- colnames(a)

# driver qual table
rbind(a,b)


# DQ pie chart

library(scales)

jj.keller %>% 
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

jj.keller %>% 
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
  
  #geom_text(aes(y = freq/2 + c(0, cumsum(freq)[-length(freq)]), 
  #              label = percent(freq/100)), size=5) +
  guides(fill = guide_legend(override.aes = list(colour=NA)))
