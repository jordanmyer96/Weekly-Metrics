library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)

setwd("C:/Users/JordanMyer/Desktop/New OneDrive/Emanate Life Sciences/DM - Inflammatix - Documents/INF-04/11. Clinical Progamming/11.3 Production Reports/11.3.3 Input Files")
source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")

CRAAssignments <- read.xlsx("CRA Assignments.xlsx")
importQueries <- read.xlsx(MostRecentFile("Medrio Reports/",".*Medrio_QueryExport_LIVE_Inflammatix_INF_04.*xlsx$","ctime"))
workingQueries <- importQueries%>%
  mutate(SITE = substring(SITE,1,3))
openQueries <- workingQueries%>%
  filter(STATUS=="Open")
openWithoutAdj <- workingQueries%>%
  filter(STATUS=="Open")%>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")

#-----
msg <- msg <- "<p>Hello and happy Monday!</p>
<p>Here is your weekly query report. Let me know if you have any questions.&nbsp;</p>
<p>Best,&nbsp;</p>
<p>Jordan</p>
<p><br></p>
<p style=\'margin:0in;font-size:15px;font-family:\"Calibri\",sans-serif;\'>Jordan Myer|Clinical Data Associate</p>
<p style=\'margin:0in;font-size:15px;font-family:\"Calibri\",sans-serif;\'>Inflammatix | (831) 824-4928</p>
<p style=\'margin:0in;font-size:15px;font-family:\"Calibri\",sans-serif;\'><a href=\"mailto:jordan@emanatelifesciences.com\"><span style=\"color:#0563C1;\">jmyer@inflammatix.com</span></a></p>"

#----Open Query Alerts----
#table(openQueries%>%filter(ASSIGNED.TO!="Unassigned")%>%pull(ASSIGNED.TO))

finalCRAAssignments <- CRAAssignments %>% 
  mutate(Created.By = paste(Name," (",Email,")",sep = ""))

oldOpenUnassigned <- importQueries %>%
  filter(STATUS == "Open"&is.na(CREATED.BY)&DAY.OPEN>7&is.na(RESPONSES)) %>%
  filter(QUERY.NAME!="releaseadjudReleasedForAdjudication:Missing Data")%>% 
  mutate(SITE = substring(SITE,1,3)) %>% 
  mutate(NewAssignment = "")

for(i in 1:nrow(oldOpenUnassigned)){
  toAssign <-  finalCRAAssignments$Created.By[grep(oldOpenUnassigned$SITE[i],finalCRAAssignments$Sites)]
  pastedToAssign <- paste(toAssign,collapse = ", ")
  oldOpenUnassigned$NewAssignment[i] <- pastedToAssign
}


generateQueryAlerts <- function(userCreated){
  oneUserWorkbook <- createWorkbook()
  
  addWorksheet(oneUserWorkbook,"Answered Queries")
  addWorksheet(oneUserWorkbook,"Unanswered Queries - Manual")
  addWorksheet(oneUserWorkbook,"Unanswered Queries - System")
  

  answeredQueriesDF <- importQueries %>% 
    filter(STATUS=="Open"&!is.na(RESPONSES)) %>% 
    filter(CREATED.BY==userCreated) %>% 
    select(c(3,5,6,7,22,21,8,12,17,20)) %>% 
    arrange(desc(DAY.OPEN))
  
  writeDataTable(oneUserWorkbook,"Answered Queries",answeredQueriesDF)
  setColWidths(oneUserWorkbook,1,1:ncol(answeredQueriesDF),widths ="auto")
  
  
  oldUnansweredQueriesDF <- importQueries %>% 
    filter(STATUS=="Open"&is.na(RESPONSES)&DAY.OPEN>7) %>% 
    select(c(3,5,6,7,22,21,8,12,17,22,20)) %>% 
    arrange(desc(DAY.OPEN))
  
  writeDataTable(oneUserWorkbook,2,oldUnansweredQueriesDF)
  setColWidths(oneUserWorkbook,2,1:ncol(oldUnansweredQueriesDF),widths ="auto")
  
  oldOpenUnassignedQueriesDF <- oldOpenUnassigned %>% 
    filter(grepl(gsub(".*\\((.*)\\).*", "\\1", userCreated),NewAssignment)) %>% 
    select(c(3,5,6,7,22,21,8,12,17,22)) %>% 
    arrange(desc(DAY.OPEN))
  
  writeDataTable(oneUserWorkbook,3,oldOpenUnassignedQueriesDF)
  setColWidths(oneUserWorkbook,3,1:ncol(oldUnansweredQueriesDF),widths ="auto")
  
  wbName <- paste("../11.3.4 Output Files/Outgoing Weekly Reports/Query Reports/",gsub(".*\\((.*)\\).*", "\\1", userCreated),".xlsx",sep = "")
  saveWorkbook(oneUserWorkbook,wbName,overwrite = TRUE)
  
  emailAddress <- gsub(".*\\((.*)\\).*", "\\1", userCreated)
  
  
}

unique(finalCRAAssignments$Created.By)
allCreatedBy <- unique(finalCRAAssignments$Created.By)

tictoc::tic()
for(j in 1:length(allCreatedBy)){generateQueryAlerts(allCreatedBy[j])}
tictoc::toc()