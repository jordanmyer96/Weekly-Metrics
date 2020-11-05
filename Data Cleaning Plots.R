# ---- Program Header ----
# Project: 
# Program Name: Expected vs Entered
# Author: 
# Created: 
# Purpose: 
# Revision History:
# Date        Author        Revision
# 

# ---- Initialize Libraries ----
library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)
library(ggplot2)
library(gridExtra)
# ---- Load Functions ----
setwd("C:/Users/JordanMyer/Desktop/New OneDrive/Emanate Life Sciences/DM - Inflammatix - Documents/INF-04/11. Clinical Progamming/11.3 Production Reports/11.3.3 Input Files")
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")

# ---- Load Raw Data ----
dataSummaryReport <- read.xlsx(MostRecentFile("Data Entry Report/",".*Medrio_SiteDataSummaryReport.*xlsx$","ctime"))
#dataSummaryReport <- read.xlsx("Data Entry Report/Medrio_SiteDataSummaryReport.xlsx")

# ---- Apply Data Transformations ----
trimSummary <- dataSummaryReport %>% 
  filter(Site == "") %>% 
  mutate(Site = substring(Subject,1,nchar(Subject)-6)) %>% 
  mutate_at(-c(1,2,3),as.numeric) %>% 
  mutate(Site = substring(Site,1,3)) %>% 
  select(-c(2,3))

dmRevCol <- c(105,35,0,759,0,2,231,0,36,54,260,sum(105,35,0,759,0,2,231,0,36,54,260))
trimSummary$Form.DM.Approved <- dmRevCol

withPercents <- trimSummary %>% 
  mutate(`% Completed (of Expected)` = round(100*Forms.Complete/(Forms.Complete+Forms.Not.Complete),0)) %>% 
  mutate(`% SDV Completed (of Entered)` = round(100*Form.Monitored/Forms.Complete)) %>% 
  mutate(`% DM Reviewed (of Entered)` = round(100*Form.DM.Approved/Forms.Complete)) %>% 
  filter(Site!="")


# ---- Generate Output ----

completedPlot <- ggplot(withPercents,aes(x = Site, y = `% Completed (of Expected)`))+
  geom_col(fill = "dark red")+coord_flip(ylim = c(0,100))
completedPlot <- completedPlot+labs(y = "Percent (%)",x = "Site")
completedPlot <- completedPlot+ggtitle("% Completed (of Expected)")
completedPlot <- completedPlot+theme(plot.title = element_text(size = 15, hjust = .5),
                       axis.title.x = element_text(face = "plain",size = 12),
                       axis.title.y = element_text(face = "plain",size = 12))+
  
  geom_text(aes(label=`% Completed (of Expected)`), position=position_dodge(width=0.9),vjust=.5, hjust=-0.8)
completedPlot
ggsave("../11.3.4 Output Files/Weekly Meeting Metrics/Percent Complete Plot.png",
       plot = completedPlot,width = 195,height = 106,units = "mm")


SDVPlot <- ggplot(withPercents,aes(x = Site, y = `% SDV Completed (of Entered)`))+
  geom_col(fill = "dark red")+coord_flip(ylim = c(0,100))
SDVPlot <- SDVPlot+labs(y = "Percent (%)",x = "Site")
SDVPlot <- SDVPlot+ggtitle("% SDV Completed (of Entered)")
SDVPlot <- SDVPlot+theme(plot.title = element_text(size = 15, hjust = .5),
                                     axis.title.x = element_text(face = "plain",size = 12),
                                     axis.title.y = element_text(face = "plain",size = 12))+
  
  geom_text(aes(label=`% SDV Completed (of Entered)`), position=position_dodge(width=0.9),vjust=.5, hjust=-0.8)

ggsave("../11.3.4 Output Files/Weekly Meeting Metrics/Percent SDV Plot.png",
       plot = SDVPlot,width = 195,height = 106,units = "mm")

DMRevPlot <- ggplot(withPercents,aes(x = Site, y = `% DM Reviewed (of Entered)`))+
  geom_col(fill = "dark red")+coord_flip(ylim = c(0,100))
DMRevPlot <- DMRevPlot+labs(y = "Percent (%)",x = "Site")
DMRevPlot <- DMRevPlot+ggtitle("% DM Reviewed (of Entered)")
DMRevPlot <- DMRevPlot+theme(plot.title = element_text(size = 15, hjust = .5),
                         axis.title.x = element_text(face = "plain",size = 12),
                         axis.title.y = element_text(face = "plain",size = 12))+
  
  geom_text(aes(label=`% DM Reviewed (of Entered)`), position=position_dodge(width=0.9),vjust=.5, hjust=-0.8)

ggsave("../11.3.4 Output Files/Weekly Meeting Metrics/Percent DM Rev Plot.png",
       plot = DMRevPlot,width = 195,height = 106,units = "mm")


#Final Table


finalTable <- trimSummary %>% 
  mutate(`% Completed (of Expected)` = round(100*Forms.Complete/(Forms.Complete+Forms.Not.Complete),0)) %>% 
  mutate(`% SDV Completed (of Entered)` = round(100*Form.Monitored/Forms.Complete)) %>% 
  mutate(`% DM Reviewed (of Entered)` = round(100*Form.DM.Approved/Forms.Complete)) %>% 
  mutate(Forms.Expected = Forms.Complete+Forms.Not.Complete) %>% 
  select(c(1,14,2,3,5,6,7,11,12,13)) %>% 
  set_names(c("Site","Entered","Complete","Expected","Overdue",
              "Monitored","DM Approved","% Completed (of Expected)",
              "% SDV Completed (of Entered)","% DM Reviewed (of Entered)")) %>% 
  arrange(desc(Site))

finalTable$Site[12] <- "Total"
dataCleaningWB<-  createWorkbook()
addWorksheet(dataCleaningWB,sheetName = "Data Cleaning Table")
writeDataTable(dataCleaningWB,sheet = 1, finalTable)
setColWidths(dataCleaningWB,sheet = 1, cols = 1:ncol(finalTable),widths = "auto")
saveWorkbook(dataCleaningWB,paste("../11.3.4 Output Files/Weekly Meeting Metrics/Data Cleaning Table ",Sys.Date(),".xlsx",sep = ""),overwrite = TRUE)



