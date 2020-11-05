# ---- Program Header ----
# Project: 
# Program Name: 
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

# ---- Load Functions ----
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")

# ---- Load Raw Data ----
importDOV <- read_sas("EDC/SV.sas7bdat")
importIC <- read_sas("EDC/DS_IC.sas7bdat")
importDem <- read_sas("EDC/DM.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")
importEW <- read_sas("EDC/DS_EW.sas7bdat")
importFUPhone <- read_sas("EDC/FU.sas7bdat")
importIE <- read_sas("EDC/IE.sas7bdat")
importIM <- read_sas("EDC/IM.sas7bdat")
importLBCOLL <- read_sas("EDC/LB_COLL.sas7bdat")
importLBSITE<- read_sas("EDC/LB_SITE.sas7bdat")
importMBSITE<- read_sas("EDC/MB_SITE.sas7bdat")
importMH<- read_sas("EDC/MH.sas7bdat")
importSOFA<- read_sas("EDC/QS_SOFA.sas7bdat")
importTOE <- read_sas("EDC/QS_TOE.sas7bdat")
importVS <- read_sas("EDC/VS.sas7bdat")


# ---- Apply Data Transformations ----



visitDateTable <- importDOV%>%
  mutate(SubjAndVisit = paste(SubjectID, Visit)) %>% 
  select(c(14,7))%>%
  set_names(c("SubjAndVisit","Visit Date"))

getEntryTime <- function(form,formString){
  columnCount <- ncol(form)
  createCol <- columnCount-2
  
  form$SubjAndVisit <- ""
  form$DateOfVisit <- as.Date("2020-01-01")
  form$Difference <- ""
  form$SourceForm <- formString
for(i in 1:nrow(form)){
  form$SubjAndVisit[i] <- paste(form[i,2],form[i,5])
  form$DateOfVisit[i] <- visitDateTable$`Visit Date`[match(form[[i,columnCount+1]],visitDateTable$SubjAndVisit)]
  form$Difference[i] <- as.Date(form[[i,columnCount+2]])-as.Date(form[[i,createCol]])
  form$Difference[i] <- as.numeric(form$Difference[i])*-1
  form$Difference[i] <- ifelse(test = form$Difference[i]<0, yes = 0, as.numeric(form$Difference[i]))
  
  }
 
return(form %>% select(c(2,3,5,ncol(form)-1,ncol(form))))}


differentGetEntryTime <- function(form,dateCol,formString){
  columnCount <- ncol(form)
  createCol <- columnCount-2
  
  form$Difference <- ""
  form$SourceForm <- formString
  form <- form %>% 
    filter(!is.na(form[,dateCol]))
  
  for(i in 1:nrow(form)){
    
    form$Difference[i] <- as.Date(form[[i,dateCol]])-as.Date(form[[i,createCol]])
    form$Difference[i] <- as.numeric(form$Difference[i])*-1
    form$Difference[i] <- ifelse(test = form$Difference[i]<0, yes = 0, as.numeric(form$Difference[i]))
  }
  
  return(form %>% select(c(2,3,5,ncol(form)-1,ncol(form))))}



newDOV <- getEntryTime(importDOV,"importDOV")
newIC <- getEntryTime(importIC,"importIC")
newDem <- getEntryTime(importDem,"importDem")
newIE <- getEntryTime(importIE,"importIE")
newIM <- getEntryTime(importIM,"importIM")
newLBCOLL <- getEntryTime(importLBCOLL,"importLBCOLL")
newLBSITE<- getEntryTime(importLBSITE,"importLBSITE")
newMBSITE<- getEntryTime(importMBSITE,"importMBSITE")
newMH<- getEntryTime(importMH,"importMH")
newSOFA<- getEntryTime(importSOFA,"importSOFA")
newTOE <- getEntryTime(importTOE,"importTOE")
newVS <- getEntryTime(importVS,"importVS")



differentNewSD <- differentGetEntryTime(importSD,9,"importSD")
differentNewEW <- differentGetEntryTime(importEW,7,"importEW")
differentNewFUPhone <- differentGetEntryTime(importFUPhone,13,"importFUPhone")



allVisitsTogether <- rbind(newDem, newDOV, newIC, newIE, newIM, 
      newLBCOLL, newLBSITE,  newMBSITE, newMH, newSOFA, 
      newTOE, newVS,differentNewFUPhone, differentNewSD, differentNewEW
      ) %>% mutate(Difference = as.numeric(Difference))

class(allVisitsTogether$Difference[1])
view(allVisitsTogether)

averageBySite <- group_by(allVisitsTogether, Site) %>% summarize(`Average Days Before Entry` = round(mean(Difference),0))
mean(averageBySite$`Average Days Before Entry`)
view(averageBySite)

medianBySite <- group_by(allVisitsTogether, Site) %>% summarize(`Average Days Before Entry` = round(median(Difference),0))


entryTimeWB<-  createWorkbook()
addWorksheet(entryTimeWB,sheetName = "Average Entry Time")
writeDataTable(entryTimeWB,sheet = 1, averageBySite)
setColWidths(entryTimeWB,sheet = 1, cols = 1:ncol(averageBySite),widths = "auto")
saveWorkbook(entryTimeWB,paste("../11.3.4 Output Files/Weekly Meeting Metrics/Average Entry Time ",Sys.Date(),".xlsx",sep = ""),overwrite = TRUE)



###For Meeting###
allVisitsTogether <- allVisitsTogether %>% 
  mutate(Difference = as.numeric(Difference))

allVisitsWithoutV0 <- allVisitsTogether %>% filter(Visit != "Day 0")
averageBySite <- group_by(allVisitsWithoutV0, Site) %>% summarize(`Average Days Before Entry` = round(mean(Difference),0))

unique(allVisitsTogether$Site)
site113NoV0 <- allVisitsTogether %>% filter(Site =="113 - Mayo Clinic - Rochester") %>% filter(Visit!="Day 0")
mean(site113NoV0$Difference)
mean(allVisitsTogether %>% filter(Site =="106 - University of Iowa") %>% filter(Visit!="Day 0") %>% pull(Difference))
mean(allVisitsTogether %>% filter(Site =="102 - Texas Tech University Health Sciences Center - El Paso") %>% filter(Visit!="Day 0") %>% pull(Difference))
mean(allVisitsTogether %>% filter(Site =="113 - Mayo Clinic - Rochester") %>% filter(Visit!="Day 0") %>% pull(Difference))
mean(allVisitsTogether %>% filter(Site =="119 - University of California Davis Medical Center" ) %>% filter(Visit!="Day 0") %>% pull(Difference))

mean(allVisitsTogether %>% filter(Visit!="Day 0") %>% pull(Difference))


# ---- Generate Output ----
#Arrange columns for final output


#Export to Excel


#Export R Object


#Generate DTD


# ---- Send Notifications ----
#Email output to recipients


# ---- Clean Environment ----

