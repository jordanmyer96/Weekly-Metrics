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
importVisitDT <- read_sas("EDC/SV.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")

enrolledSubjects <- unique(importVisitDT %>% filter(SubjectStatus == "Enrolled") %>%  pull(SubjectID))


workingQueries <- importQueries%>%
  filter(SUBJECT.ID%in%enrolledSubjects)

allOpenQueries <- workingQueries%>%
  filter(STATUS=="Open")%>%
  filter(!is.na(CREATED.BY)) %>% 
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

openWithNoAnswer <- allOpenQueries %>% 
  filter(is.na(RESPONSES))

openWithAnswer <- allOpenQueries %>% 
  filter(!is.na(RESPONSES))

closedQueries <- workingQueries%>%
  filter(STATUS!="Open")%>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

completedStudy <- importSD %>% filter(DSDECOD=="28 day data collection complete")
#---- Unassigned Listing----

#-----Open Vs Answered


totalOpenQueries <- nrow(allOpenQueries)
numberOfUnanswered <- nrow(openWithNoAnswer)
averageDaysUnanswered <- round(mean(openWithNoAnswer%>% pull(DAY.OPEN)),0)

numberOfAnswered <- nrow(openWithAnswer)
averageDaysAnswered <- round(mean(openWithAnswer%>% pull(DAY.OPEN)),0)
totalAverageDays <- round(mean(allOpenQueries$DAY.OPEN),0)

openVAnsweredDF <- data.frame(Type = c("Open Without Response","Answered not Closed","Total Open Queries"),
                              Count = c(numberOfUnanswered,numberOfAnswered,totalOpenQueries),
                              `Average Age` = c(averageDaysUnanswered,averageDaysAnswered,totalAverageDays))
openVAnsweredDF

#-----Query Aging-----
lessThan7 <- allOpenQueries%>%filter(DAY.OPEN<7)
sevenTo21 <- allOpenQueries%>%filter(between(DAY.OPEN,7,21))
moreThan21 <- allOpenQueries%>%filter(DAY.OPEN>21)
totalQueries <- sum(nrow(lessThan7),nrow(sevenTo21),nrow(moreThan21))
queryAgingDF <- data.frame(QueryAge=c("<7 Days","7-21 Days",">21","Total"),
                           Count=c(nrow(lessThan7),nrow(sevenTo21),nrow(moreThan21),totalQueries))
view(queryAgingDF)


df <-  openWithNoAnswer
agingBySite <- function(df){
lessThan7 <- df%>%filter(DAY.OPEN<7)
sevenTo21 <- df%>%filter(between(DAY.OPEN,7,21))
moreThan21 <- df%>%filter(DAY.OPEN>21)

lessThan7BySite <- group_by(lessThan7, SITE) %>% 
  summarize(`Less Than 7 By Site` = round(length(DAY.OPEN),0))

sevenTo21BySite <- group_by(sevenTo21, SITE) %>% 
  summarize(`Seven to 21 By Site` = round(length(DAY.OPEN),0))

moreThan21BySite <- group_by(moreThan21, SITE) %>% 
  summarize(`More Than 21 By Site` = round(length(DAY.OPEN),0))


allTogetherBySite <- full_join(full_join(lessThan7BySite,sevenTo21BySite),moreThan21BySite) %>% 
  mutate_all(~replace(., is.na(.), 0)) %>% 
  rbind(totalsBySite <- c("Total",sum(lessThan7BySite[,2]),sum(sevenTo21BySite[,2]),sum(moreThan21BySite[,2])))

return(allTogetherBySite)}

agingBySite(openWithNoAnswer)

