library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)

setwd("C:/Users/JordanMyer/Desktop/New OneDrive/Emanate Life Sciences/DM - Inflammatix - Documents/INF-04/11. Clinical Progamming/11.3 Production Reports/11.3.3 Input Files")
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")


importQueries <- read.xlsx(MostRecentFile("Medrio Reports/",".*Medrio_QueryExport_LIVE_Inflammatix_INF_04.*xlsx$","ctime"))

openQueries <- importQueries%>%
  filter(STATUS=="Open")
openWithoutAdj <- importQueries%>%
  filter(STATUS=="Open")%>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")



#----Open Query Alerts----
#table(openQueries%>%filter(ASSIGNED.TO!="Unassigned")%>%pull(ASSIGNED.TO))

wb <- createWorkbook()

addWorksheet(wb,"Answered Queries")
addWorksheet(wb,"Unanswered Queries")
addWorksheet(wb,"Missing Data")


answeredQueriesDF <- importQueries %>% 
  filter(STATUS=="Open"&!is.na(RESPONSES)&ASSIGNED.TO!="Unassigned") %>% 
  select(c(3,5,6,7,22,21,8,12,17,20)) %>% 
  arrange(desc(DAY.OPEN))

writeDataTable(wb,"Answered Queries",answeredQueriesDF)
setColWidths(wb,1,1:ncol(answeredQueriesDF),widths ="auto")


unansweredQueriesDF <- openWithoutAdj %>% 
  filter(is.na(RESPONSES)&DAY.OPEN>7&QUERY.TYPE!="Missing Data") %>% 
  select(3,5,7,22,21,8,12,13,17,20) %>% 
  arrange(desc(DAY.OPEN))

writeDataTable(wb,2,unansweredQueriesDF)
setColWidths(wb,2,1:ncol(unansweredQueriesDF),widths ="auto")
addStyle()

missingData <- openWithoutAdj %>% 
  filter(is.na(RESPONSES)&DAY.OPEN>7&QUERY.TYPE=="Missing Data") %>% 
  select(3,5,6,7,22,21,8,12,17) %>% 
  arrange(desc(DAY.OPEN)) 
  
writeDataTable(wb,3,missingData)
setColWidths(wb,3,1:ncol(missingData),widths ="auto")  

wbName <- paste("../11.3.4 Output Files/Outgoing Weekly Reports/Query Reports/All Subject Queries.xlsx",sep = "")
saveWorkbook(wb,wbName,overwrite = TRUE)
