library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)
library(Hmisc)
library("miscTools")

setwd("C:/Users/JordanMyer/Desktop/New OneDrive/Emanate Life Sciences/DM - Inflammatix - Documents/INF-04/11. Clinical Progamming/11.3 Production Reports/11.3.3 Input Files")
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")

importVisitDT <- read_sas("EDC/SV.sas7bdat")
importDem <- read_sas("EDC/DM.sas7bdat")
importTOE <- read_sas("EDC/QS_TOE.sas7bdat")
importIE <- read_sas("EDC/IE.sas7bdat")
importSD <- read_sas("EDC/DS_SD.sas7bdat")
importLOS <- read_sas("EDC/HO.sas7bdat")
monitoringWB <- createWorkbook()


trimDem <- importDem%>%
  select(c(2:4,7,9,11,13,15))%>%
  set_names(c("SubjectID","Site","Status","Age","Sex","Ethnicity","Race","Race - Other"))


trimIE <- importIE%>%
  select(2:4,7,9,11,13)%>%
  set_names(c("SubjectID","Site","Status","Over 18 Years Old",
              "Suspected acute infection (e.g., respiratory, urinary, abdominal, skin & soft-tissue, meningitis/encephalitis, or any other infection), and at least one of the signs in Table 1",
              "Suspected sepsis of any cause as defined by a blood culture order by the treating physician, and at least two of the signs in Table 1",
              "Able to provide informed consent, or consent by legally authorized representative"))

trimTOE <- importTOE%>%
  select(2:4,7,9)%>%
  set_names(c("SubjectID","Site","Status",
              "Suspected Primary Infection Source","Other Primary Infection Source"))


trimLOS <- importLOS %>% 
  select(2:4,9)%>%
  set_names(c("SubjectID","Site","Status",
              "Patient was admitted to"))



trimSD <- importSD%>%
  select(2:4,11,19,49,51)%>%
  set_names(c("SubjectID","Site","Status","Date of Death","Discharge Date","Readmitted","Did subject receive ICU level care during readmission?"))%>%
  #mutate(TestYN = `Date of Death`-`Discharge Date`)
  mutate(`Death Within 30d of Discharge` = ifelse(test = !is.na(`Date of Death`)&(`Date of Death`-`Discharge Date`)<30,yes = "Yes",ifelse(test = !is.na(`Date of Death`),yes = "No",no = "")))%>%
  mutate(`Death Within 30d of Discharge` = ifelse(test = is.na(`Death Within 30d of Discharge`),yes = "Yes",no = `Death Within 30d of Discharge`))%>%
  mutate(`Death Within 30d of Discharge` = ifelse(test = !is.na(`Date of Death`)&grepl("Yes",Readmitted),yes = "Death after Readmission",no = `Death Within 30d of Discharge`))%>%
  select(1:3,7:8)



allTogether <- left_join(left_join(left_join(left_join(trimDem,trimIE),trimTOE),trimSD),trimLOS)%>%
  filter(Status=="Enrolled")%>%
  mutate(`Did subject receive ICU level care?` = ifelse(test = `Did subject receive ICU level care during readmission?`=="Yes",yes = "Yes - Readmission", ifelse(test = grepl("ICU",`Patient was admitted to`), yes = "Yes", no = "No"))) %>% 
  select(c(-15,-17))



infectionTypeTable <- as.data.frame(table(allTogether$`Suspected Primary Infection Source`))%>%
  arrange(desc(Freq))%>%
  set_names(c("Infection Type","Count"))



addWorksheet(monitoringWB,"Summary Table")
desiredColumns <- c(2,5:16)
startColNumber <- 1
dfVector <- c()
for(j in 1:ncol(allTogether)){
  if(j%in%desiredColumns){
    for(i in 1:length(names(table(allTogether[,j])))){
      dfVector[2*i-1] <- names(table(allTogether[,j]))[i]
      dfVector[2*i] <- table(allTogether[,j])[i]
      toRemove <- c(which(""%in%dfVector),1+which(""%in%dfVector))
      ifelse(test = length(toRemove)>0, 
             yes = finalCol <- data.frame(Name = dfVector[-toRemove]), 
             no = finalCol <- data.frame(Name = dfVector))
      
    }
    colnames(finalCol) <- names(allTogether)[j]
    writeDataTable(monitoringWB,sheet = 1,finalCol,startCol = startColNumber)
    dfVector <- c()
    toRemove <- c()
    finalCol <- c()
    cleanCol <- c()
    startColNumber <- startColNumber+1
  }
}

setColWidths(monitoringWB,sheet = 1,cols = c(1:13), widths = c(54,20,24,29,20,24,55,45,35,45,32,35,54))
headerStyle <- createStyle(wrapText = TRUE)
addStyle(monitoringWB,sheet = 1, headerStyle,1,cols = 1:13)

addWorksheet(monitoringWB,"Subject Lines")
writeDataTable(monitoringWB,sheet = 2, allTogether,withFilter = TRUE)
setColWidths(monitoringWB,sheet = 2,cols = c(1,2,8,13,14,15,16), widths = c(10.4,49.7,18.5,32,32,52,34))

addWorksheet(monitoringWB,"Type of Infection")
writeDataTable(monitoringWB,sheet = 3, infectionTypeTable)
setColWidths(monitoringWB,sheet = 3,cols = 1, widths = 22.7)

saveWorkbook(monitoringWB,paste("../11.3.4 Output Files/Weekly Meeting Metrics/Monitoring Report ",Sys.time()%>%format('%b%d%Y'),".xlsx",sep=""),overwrite = TRUE)

