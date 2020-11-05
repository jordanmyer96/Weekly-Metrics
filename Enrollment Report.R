library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(tidyverse)
library(janitor)

source("../11.3.1  R Production Programs/INF Global Functions.R")
cleaner()
source("../11.3.1  R Production Programs/INF Global Functions.R")

importVisitDT <- read_sas("EDC/SV.sas7bdat")

enrollmentDates <- importVisitDT %>% filter(SubjectStatus!="Screen Failure") %>% 
  filter(Visit == "Day 0") %>% 
  mutate(Month = month(SVDAT,label = TRUE))

enrollmentDF<- as.data.frame.matrix(table(enrollmentDates$Site,enrollmentDates$Month),stringsAsFactors = FALSE)


enrollmentDF$Site <- c(rownames(table(enrollmentDates$Site,enrollmentDates$Month)))
enrollmentDF <- enrollmentDF %>% mutate(Site = substring(Site,1,3))
enrollmentDF <- enrollmentDF %>% select(13,1:12)
finalEnrollmentDF <- enrollmentDF %>% 
  adorn_totals(where = "col") %>% 
  adorn_totals(where = "row")

enrollmentWB <- createWorkbook()
addWorksheet(enrollmentWB,"Enrollment Table")
writeDataTable(enrollmentWB,sheet = 1, 
               finalEnrollmentDF,
               tableStyle = "TableStyleMedium3",
               withFilter = FALSE)
setColWidths(enrollmentWB,sheet = 1, cols = 1,widths = "auto")

saveWorkbook(enrollmentWB,paste("../11.3.4 Output Files/Weekly Meeting Metrics/Enrollment Table ",Sys.Date(),".xlsx",sep = ""),overwrite = TRUE)
