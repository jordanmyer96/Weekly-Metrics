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

getEntryTime <- function(form){
  columnCount <- ncol(form)
  createCol <- columnCount-2
  
  form$SubjAndVisit <- ""
  form$DateOfVisit <- as.Date("2020-01-01")
  form$Difference <- ""
for(i in 1:nrow(form)){
  form$SubjAndVisit[i] <- paste(form[i,2],form[i,5])
  
  form$DateOfVisit[i] <- visitDateTable$`Visit Date`[match(form[[i,columnCount+1]],visitDateTable$SubjAndVisit)]
  form$Difference[i] <- as.Date(form[[i,columnCount+2]])-as.Date(form[[i,createCol]])
  form$Difference[i] <- as.numeric(form$Difference[i])*-1
  form$Difference[i] <- ifelse(test = form$Difference[i]<0, yes = 0, as.numeric(form$Difference[i]))
  }
 
return(form %>% select(c(2,3,5,ncol(form))))}

differentGetEntryTime <- function(form,dateCol){
  columnCount <- ncol(form)
  createCol <- columnCount-2
  

  form$Difference <- ""
  form <- form %>% 
    filter(!is.na(form[,dateCol]))
  
  for(i in 1:nrow(form)){
    
    form$Difference[i] <- as.Date(form[[i,dateCol]])-as.Date(form[[i,createCol]])
    form$Difference[i] <- as.numeric(form$Difference[i])*-1
    form$Difference[i] <- ifelse(test = form$Difference[i]<0, yes = 0, as.numeric(form$Difference[i]))
  }
  
  return(form %>% select(c(2,3,5,ncol(form))))}



newDOV <- getEntryTime(importDOV)
newIC <- getEntryTime(importIC)
newDem <- getEntryTime(importDem)
newIE <- getEntryTime(importIE)
newIM <- getEntryTime(importIM)
newLBCOLL <- getEntryTime(importLBCOLL)
newLBSITE<- getEntryTime(importLBSITE)
newMBSITE<- getEntryTime(importMBSITE)
newMH<- getEntryTime(importMH)
newSOFA<- getEntryTime(importSOFA)
newTOE <- getEntryTime(importTOE)
newVS <- getEntryTime(importVS)

differentNewSD <- differentGetEntryTime(importSD,9)
differentNewEW <- differentGetEntryTime(importEW,7)
differentNewFUPhone <- differentGetEntryTime(importFUPhone,13)



allVisitsTogether <- rbind(newDem, newDOV, newIC, newIE, newIM, 
      newLBCOLL, newLBSITE,  newMBSITE, newMH, newSOFA, 
      newTOE, newVS,differentNewFUPhone, differentNewSD, differentNewEW
      )

allVisitsTogether <- allVisitsTogether %>% 
  mutate(Difference = as.numeric(Difference))

averageBySite <- group_by(allVisitsTogether, Site) %>% summarize(`Average Days Before Entry` = round(mean(Difference),0))

view(averageBySite)

entryTimeWB<-  createWorkbook()
addWorksheet(entryTimeWB,sheetName = "Average Entry Time")
writeDataTable(entryTimeWB,sheet = 1, averageBySite)
setColWidths(entryTimeWB,sheet = 1, cols = 1:ncol(averageBySite),widths = "auto")
saveWorkbook(entryTimeWB,paste("../11.3.4 Output Files/Weekly Meeting Metrics/Average Entry Time ",Sys.Date(),".xlsx",sep = ""))

# ---- Generate Output ----
#Arrange columns for final output


#Export to Excel


#Export R Object


#Generate DTD


# ---- Send Notifications ----
#Email output to recipients


# ---- Clean Environment ----

