
  form <- importDOV
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
  