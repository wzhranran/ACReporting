#' @name Scoring
#' @title A function to trim and score responses.
#' @description  A function to score the responses, compute percentage correct, and classify based on it.
#' Reading is always analyzed before Listening.
#' @param MainPath The folder where the response files are stored in
#' @param Language Testing language
#' @param TestName Name of the test
#' @param AdminDate Date of test administration
#' @param ProfScale Proficiency scale, "ACTFL" or "ILR"
#' @param ProfVersion_l The version of the proficiency level file downloaded from ACTesting for listening
#' @param ProfVersion_r The version of the proficiency level file downloaded from ACTesting for reading
#' @param n_demo number of demographic variables, usually it's 6.
#' @param PL_file logical. The default is FALSE, PL file is generated from Proficiency file downloaded from ACTesting. If TRUE, additional PL file is provided, usually after IDR.
#' After IDR, just change the last column of PL file, including changing proficiency level and changing to DNS.
#' @param keychange_r The default is NULL. If the reading key is changed from IDR, a list of item index and new keys should be here.
#' The first element is item index, the rest are the keys. Each item with key change is an object in the list.
#' @param keychange_l The default is NULL. If the listening key is changed from IDR, a list of item index and new keys should be here.
#' @return All the files are stored in the determined folders. The excel file is also generated.
#' @examples
#' Scoring(MainPath, Language, TestName, AdminDate, ProfScale, ProfVersion_l, ProfVerson_r, n_demo, PL_file=F,
#' keychange_r=list(c(1,1,2), c(13,1,4)), keychange_l=NULL)

Scoring<-function(MainPath, Language, TestName, AdminDate, ProfScale, ProfVersion_l, ProfVersion_r, n_demo=6, PL_file=F,
                  keychange_r=NULL, keychange_l=NULL)
{
  library(openxlsx)
  #### Reading ####
  Skill<-"Reading"
  # Read in response file
  response<-ReadCSVFile(paste0(MainPath,"/"),paste0(TestName, "_", Skill))
  # Put response file in Data/clean and Data/original
  CleanPath<-paste0(MainPath,"/Data/clean/")
  if (file.exists(CleanPath)==F)
  {
    if (file.exists(paste0(MainPath,"/Data/"))==F)
    {
      dir.create(paste0(MainPath,"/Data/"))
      dir.create(CleanPath)
    } else dir.create(CleanPath)
  }
  write.csv(response, paste0(CleanPath,paste0(TestName, "_", Skill,".csv")),row.names = F)
  OriginalPath<-paste0(MainPath,"/Data/original/")
  if (file.exists(OriginalPath)==F) dir.create(OriginalPath)
  write.csv(response, paste0(OriginalPath,paste0(TestName, "_", Skill,".csv")),row.names = F)

  PerNo<-nrow(response) #number of examinees
  ItemNo<-ncol(response)-n_demo
  RandomItemIDs<-colnames(response)[-c(1:n_demo)] #unique item ID

  if (PL_file==F)
  {
    #### Read in proficiency level file form ACTesting ####
    OriPLs<-ReadCSVFile(paste0(MainPath,"/"),paste(Language, Skill, ProfVersion_r, "levels", sep=" "))
    write.csv(OriPLs, paste0(CleanPath,paste(Language, Skill, ProfVersion_r, "levels", sep=" "),".csv"),row.names = F)
    PL<-matrix(,ncol=4, nrow=ItemNo)
    word<-ifelse(Skill=="Reading","read_","listen_")
    PL[,1]<-c(paste0(word,"0",1:9),paste0(word,10:ItemNo))
    if (ProfScale=="ACTFL")
    {
      PL[,-1]<-as.character(OriPLs$ACTFL)
    } else PL[,-1]<-as.character(OriPLs$ILR)
    colnames(PL)<-c("ItemNo","Original_Prof_Level","Final_Prof_Level","Reported_Prof_Level")
    ItemPLsFile <- paste(TestName,"_",Skill,"_PLs",sep = "")
    write.csv(PL, paste0(CleanPath,ItemPLsFile, ".csv"), row.names = F)
  } else {
    ItemPLsFile <- paste(TestName,"_",Skill,"_PLs",sep = "")
    PL<-ReadCSVFile(CleanPath, ItemPLsFile)
    PL<-as.matrix(PL)
  }

  #### Scoring and assign proficiency levels ####
  if (ProfScale == "ACTFL"){
    # allowed values for ACTFL proficiency scale.
    allowedValues <- c("NL", "NM", "NH", "NL/NM", "NM/NH", "Novice", "NH/IL",
                       "IL", "IM", "IH", "IL/IM", "IM/IH", "Intermediate", "IH/AL",
                       "AL", "AM", "AH", "AL/AM", "AM/AH", "Advanced", "S", "D",
                       "DNS")
  } else if(ProfScale == "ILR"){
    ## allowed values for ILR proficiency scale.
    allowedValues <- c("0", "0+", "0/0+", "0+/1", "1", "1+", "1/1+", "1+/2",
                       "2", "2+", "2/2+", "2+/3", "3", "3+", "3/3+", "3+/4",
                       "4", "4+", "4/4+", "4+/5", "5", "DNS")
  } else{
    # Stop the program if proficiency scale is not "ACTFL" or "ILR."
    stop(paste("Invalid Proficiency Scale. Score calculations were stopped."))
  }

  c <- unique(PL[,4]) %in% allowedValues

  # Check if the values in the file with test proficiency levels match the ones in allowedValues variable.
  # Stop if there is a mismatch, continue otherwise.
  if ("FALSE" %in% c)
    {
    stop (paste("Invalid values assigned to proficiency levels") )
    } else{
     ReportedLevels <- unique(PL[,4]) [unique(PL[,4]) != "DNS"]
     ReportedLevels <- ReportedLevels[order(match(ReportedLevels, allowedValues))]
     print(paste("The following are proficiency levels in",Skill, ":",
                paste(ReportedLevels, collapse = ", "), collapse = " "))
    }
  n_levels<-length(ReportedLevels)

  # Replace column names, e.g. Q1 to either read_01 or listen_01, etc.
  colnames(response)[-c(1:n_demo)] <- PL[,1]
  # Remove items which are marked as "DNS" as "Do not score" from calculations.
  if ("DNS" %in% as.character(PL[,4])){
    pl <- PL[PL[,4] == "DNS",1] # name of DNS item
    response_pure <- response[, -c(1:n_demo)] # response matrix without demographic data
    response_pure <- response_pure[, !(names(response_pure) %in% pl)]
    print(paste("Dropped from calculations - item(s):",
                paste(pl, collapse = ", "), collapse = " "))
  } else{

    response_pure <- response[, -c(1:n_demo)]
    print("All items are taken into account in the calculations.")
  }

  ## Count Total Omitted, Total Correct, and Total Correct Percent for
  ## Reading or Listening data file.

  Omitted <- apply(response_pure, 1, FUN = function(x) length(which(x=="NA"))+length(which(is.na(x)))) # -1 is omitted in csv, however NA here
  score<-response_pure
  if (length(keychange_r)!=0)
  {
    for (k in 1:length(keychange_r))
    {
      score[,keychange_r[[k]][1]]<-ifelse(response_pure[,keychange_r[[k]][1]] %in% keychange_r[[k]][-1],1,0)
    }
  }
  score[,which(PL[,4]=="DNS")]<-0
  TotalCor <- apply(score, 1, FUN = function(x) length(which(x == 1)))
  TotalCorPcnt <- apply(score, 1, FUN = function(x) (length(which(x == 1)))/ItemNo*100)
  TotalCorPcnt <-round(TotalCorPcnt,digits=2)

  ## Subset dataframe by proficiency levels and counting Number Correct,
  ## Percent Correct, mastery at 75% and above, and mastery at 67% and above
  ## with each proficiency level.

  results <- data.frame()

  for(pls in ReportedLevels){

    ## Redefine the '[' function by adding "drop = FALSE" to prevent from reducing
    ## a matrix into a vector during calculations. For example, if there is only
    ## one item in a proficiency level, the '[' automatically reduces data into
    ## a vector. With "drop = FALSE" change in the "base" package this reduction
    ## is disabled. However, the function is restored after the calculations are
    ## complete so that the change doesn't interfere with the '[' functions defined
    ## for classes other than 'matrix'.

    '[' <- function(...) base::'['(..., drop = FALSE)

    pl <- PL[PL[,4] == pls,1] # items with ProfLevel=pls
    tmp <- score[, colnames(response_pure) %in% pl] # responses belong to pls
    NumCor <- apply(tmp, 1, FUN = function(x) length(which (x == 1)))
    CorPcnt <- apply(tmp, 1, FUN = function(x) length(which(x == 1))/ncol(tmp)*100)
    CorPcnt <- round(CorPcnt,digits = 2)
    Mastery_75Up <- ifelse(CorPcnt > 74.49, "Yes", "No")
    Mastery_67Up <- ifelse(CorPcnt > 66.49, "Yes", "No")
    results <- rbind(results,
                     data.frame(NumCor, CorPcnt, Mastery_75Up, Mastery_67Up))
  }


  # Add proficiency level column for each iteration and merge that column with the results.
  # Split results dataframe into segments the number of which equals to the number of unique proficiency levels in the test section.
  # Assign proficiency level names to split dataframe level names.
  pl_75 <- ifelse(results$Mastery_75Up == "Yes",
                  as.character(rep(ReportedLevels, each = PerNo)), "")
  pl_67 <- ifelse(results$Mastery_67Up == "Yes",
                  as.character(rep(ReportedLevels, each = PerNo)), "")
  results <- cbind(results, pl_75, pl_67)
  results2 <- split(results, rep(1:n_levels,
                                 each = PerNo))
  names(results2) <- ReportedLevels
  results3 <- do.call(cbind, results2)


  # Count the number of mastery levels achieved within each proficiency level.
  data_75Up <- results3[, grep("Mastery_75|pl_75", colnames(results3))]
  Maintain_75Up <- ifelse((apply(data_75Up, 1, FUN = function(x) length(which(x == "Yes")))) == 0,
                          "None", paste("In level(s): ", apply(data_75Up[, grep("pl_75", colnames(data_75Up))],
                                 1, FUN = function(x) paste(x, collapse = ", "))))

  data_67Up <- results3[, grep("Mastery_67|pl_67", colnames(results3))]
  Maintain_67Up <- ifelse((apply(data_67Up, 1, FUN = function(x) length(which(x == "Yes")))) == 0,
                          "None", paste("In level(s): ", apply(data_67Up[, grep("pl_67", colnames(data_67Up))],
                                 1, FUN = function(x) paste(x, collapse = ", "))))

  results4 <- results3[, -c(grep("pl_", colnames(results3)))]


  ## Create a dataframe.
  ## Order by examinee's last name, first name in alphabetical order, and then by examinee ID from smallest to largest.


  colnames(response)[-c(1:6)]<-RandomItemIDs  # Return the place holder: RandomItemIDs to the data file variable names


  d <- cbind(response, Omitted, TotalCor, TotalCorPcnt, results4,
             Maintain_75Up, Maintain_67Up)
  d$Maintain_75Up <- gsub("[[:space:]],", "", d$Maintain_75Up)
  d$Maintain_75Up <- gsub(",[[:space:]]$", "", d$Maintain_75Up)
  d$Maintain_67Up <- gsub("[[:space:]],", "", d$Maintain_67Up)
  d$Maintain_67Up <- gsub(",[[:space:]]$", "", d$Maintain_67Up)

  dSorted <- d[with(d, order(d$lastName, d$firstName,d$examineeID)),]

  ## write to Scores folder
  ScoreFile<-paste0(TestName,"_",Skill,"_Mastery")
  ScorePath<-paste0(MainPath,"/Scores/")
  if (file.exists(ScorePath)==F) dir.create(ScorePath)
  write.csv(dSorted, paste(ScorePath,ScoreFile, ".csv", sep = ""),
            row.names = FALSE)

  ## Remove too many omits and repeated examinees
  Cleaned<-dSorted[!dSorted$Omitted>round(ItemNo/4*3,0),] # removing individuals missing or omited more than 3/4 of questions
  jMetrikPath<-paste0(MainPath,"/Data/jMetrik/")
  if (file.exists(jMetrikPath)==F) dir.create(jMetrikPath)
  write.csv(Cleaned[,1:(n_demo+ItemNo+3)], paste(jMetrikPath,TestName,"_",Skill,"_jMetrik.csv",sep=""), row.names = FALSE)    # write the cleaned and combined file into a csv file
  # demographic information + responses + omitted, total correct, total correct proportion

  ## The original '[' base function is being restored here.
  '[' <- function(...) base::'['(..., drop = TRUE)

    #### Create Excel File ####

    # Create style
    bold<-createStyle(textDecoration = "bold")
    border_general<-createStyle(border = "TopBottomLeftRight", borderStyle = "thin", halign = "left")
    border_head<-createStyle(border=c("Top","Bottom","Left","Right"),borderStyle = c("thick","thick","thin","thin"), halign = "left")


    wb<-createWorkbook()
    # "Reading"
    addWorksheet(wb,Skill)
    writeData(wb, Skill, TestName, startCol = 1, startRow = 1, rowNames = F, colNames=F)          # write TestName in Cell A1
    writeData(wb, Skill, paste("Language",Language,sep= ": "), startCol = 1, startRow = 2, rowNames = F, colNames=F)
    writeData(wb, Skill, paste("Skill", Skill, sep= ": "), startCol = 1, startRow = 3, rowNames = F, colNames = F)
    writeData(wb, Skill, paste(AdminDate,sep= ""), startCol = 1, startRow = 4, rowNames = F, colNames=F)         # write AdminDate(s) in Cell A4
    dSorted_reading<-dSorted
    word1<-ifelse(Skill=="Reading","R","L")
    colnames(dSorted_reading)<-c(colnames(dSorted)[1:n_demo],c(paste0(word1,"0",1:9),paste0(word1,10:ItemNo)),
                                 "Omitted Items", "#Correct", "%Correct",
                                 rep(c("#Correct", "%Correct", "75%&Up", "67%&Up"), n_levels),
                                 "Maintaining 75% and Above Correct", "Maintaining 67% and Above Correct")
    mergeCells(wb, Skill, cols=c(n_demo-1,n_demo), rows=6) # Proficiency levels
    mergeCells(wb, Skill, cols=c(n_demo+ItemNo+2,n_demo+ItemNo+3), rows=6) # Total 44 items
    head_sub<-rep(0,n_levels*4)
    for (k in 1:n_levels)
    {
      head_sub[(4*(k-1)+1):(4*k)]<-c(paste0(ReportedLevels[k]," (",length(which(PL[,4]==ReportedLevels[k]))," items)"),"","","")
      mergeCells(wb, Skill, cols=c(n_demo+ItemNo+4+(k-1)*4,n_demo+ItemNo+4*k+3),rows=6)
    }
    mergeCells(wb, Skill, cols=c(n_demo+ItemNo+4+4*n_levels, n_demo+ItemNo+5+4*n_levels), rows=6)
    head<-data.frame(t(c("Proficiency Levels", "", PL[,4], "", paste0("Total (", ItemNo, " items)"), "",
                         head_sub, paste0("Reading Comprehension (", ItemNo, " items)"))))
    writeData(wb, Skill, head, startCol = n_demo-1, startRow = 6, rowNames = F, colNames=F)
    writeData(wb, Skill, dSorted_reading, startCol = 1, startRow = 7, rowNames = F, colNames=T,keepNA = T)

    addStyle(wb, Skill, border_general, rows=8:(PerNo+7), cols=c(1:(n_demo+ItemNo+5+4*n_levels)), gridExpand =T, stack = FALSE)
    addStyle(wb, Skill, border_head, rows=c(6,7), cols=c(1:(n_demo+ItemNo+5+4*n_levels)), gridExpand = T, stack = FALSE)
    setColWidths(wb, Skill, cols=c(2:(n_demo+ItemNo+5+4*n_levels)), widths = "auto", ignoreMergedCells = T)

    # "SS_Reading"
    addWorksheet(wb,"SS_Reading")
    # index for standard setting
    ind_ss<-sort.int(dSorted$TotalCor,index.return = T)$ix
    hide<-c(grep("75Up",colnames(dSorted)),grep("67Up",colnames(dSorted)))
    ss<-dSorted_reading[ind_ss,-hide]
    writeData(wb,"SS_Reading",ss, startCol = 1, startRow = 2, rowNames = F, colNames=T,keepNA = T)
    mergeCells(wb, "SS_Reading", cols=c(n_demo+ItemNo+2, n_demo+ItemNo+3), rows=1) # merge for "Total 44 items"
    head_sub<-rep(0, 2*n_levels)
    for (k in 1:n_levels)
    {
      head_sub[(2*k-1):(2*k)]<-c(paste0(ReportedLevels[k]," (",length(which(PL[,4]==ReportedLevels[k]))," items)"),"")
      mergeCells(wb, "SS_Reading", cols=c(n_demo+ItemNo+4+(k-1)*2,n_demo+ItemNo+2*k+3),rows=1)
    }
    head<-data.frame(t(c(paste0("Total (", ItemNo, " items)"), "",head_sub)))
    writeData(wb, "SS_Reading", head, startCol= n_demo+ItemNo+2, startRow=1, rowNames = F, colNames=F)

    addStyle(wb, "SS_Reading", border_general, rows=3:(PerNo+2), cols=c(1:(n_demo+ItemNo+3+2*n_levels)), gridExpand =T, stack = FALSE)
    addStyle(wb, "SS_Reading", border_head, rows=c(1,2), cols=c(1:(n_demo+ItemNo+3+2*n_levels)), gridExpand = T, stack = FALSE)
    conditionalFormatting(wb, "SS_Reading", cols=c((n_demo+ItemNo+4):(n_demo+ItemNo+3+2*n_levels)), rows=3:(PerNo+2),
                      rule=">66", style=bold)


    # "Survey"
    addWorksheet(wb, "Survey")
    survey<-read.csv(paste0(MainPath,"/",TestName,"_survey.csv"),
                     header = TRUE, row.names = NULL,
                     flush = TRUE, fill = TRUE)
    SurveyPath<-paste0(MainPath,"/Data/survey/")
    if (file.exists(SurveyPath)==F) dir.create(SurveyPath)
    write.csv(survey,paste0(SurveyPath,TestName,"_survey.csv"),row.names = F)

    survey <- survey[with(survey, order(survey$Name, survey$Examinee_Key)),]
    writeData(wb,"Survey", survey, startCol = 1, startRow = 1, rowNames = F, colNames=T,keepNA = T)

    # "ItemAnalyses"
    addWorksheet(wb, "ItemAnalyses")
    writeFormula(wb, sheet="ItemAnalyses", x = paste0(Skill,"!A1"), startCol = "A", startRow = 1)     # write TestName in Cell A1 from "Reading" sheet into "ItemAnalyses" sheet
    writeFormula(wb, sheet="ItemAnalyses", x = paste0(Skill,"!A2"), startCol = "A", startRow = 2)     # write TestName in Cell A2 from "Reading" sheet into "ItemAnalyses" sheet
    writeFormula(wb, sheet="ItemAnalyses", x = paste0(Skill,"!A4"), startCol = "A", startRow = 3)     # write administration date(s) in Cell A4 from "Reading" sheet into "ItemAnalyses" sheet
    pl_xlsx<-cbind(PL[,c(1,2,4)],NA)
    colnames(pl_xlsx)<-c("Item No.", "Original PL", "Final PL", "Comments")
    writeData(wb, sheet="ItemAnalyses", pl_xlsx, startCol = ifelse(Skill=="Reading", "A", "F"), startRow = 6)

    levelcount<-NULL
    for (i in 1:n_levels)
    {
      levelcount[i]<-length(which(PL[,4]==ReportedLevels[i]))
    }
    levelsum<-data.frame(c(ReportedLevels,"Total"),c(levelcount,ItemNo))
    colnames(levelsum)<-c("Prof.Lvl","# of Items")
    writeFormula(wb, sheet="ItemAnalyses", x = paste0(Skill,"!A1"), startCol = "K", startRow = 1)     # write TestName in Cell A1 from "Reading" sheet into "ItemAnalyses" sheet
    writeFormula(wb, sheet="ItemAnalyses", x = paste0(Skill,"!A2"), startCol = "K", startRow = 2)
    writeFormula(wb, sheet="ItemAnalyses", x = paste0(Skill,"!A4"), startCol = "K", startRow = 3)
    writeData(wb, sheet="ItemAnalyses", Skill, startCol= ifelse(Skill=="Reading", "K", "M"), startRow = 5)# write administration date(s) in Cell A4 from "Reading" sheet into "ItemAnalyses" sheet
    writeData(wb, sheet="ItemAnalyses", levelsum, startCol = ifelse(Skill=="Reading", "K", "M"), startRow = 6)
    mergeCells(wb, "ItemAnalyses", cols = 1:4, rows=5)
    mergeCells(wb, "ItemAnalyses", cols = 6:9, rows = 5)
    writeData(wb, "ItemAnalyses", "Reading Comprehension", startCol = 1, startRow = 5)
    writeData(wb, "ItemAnalyses", "Listening Comprehension", startCol = 6, startRow = 5)

    addStyle(wb, "ItemAnalyses", border_general, rows=c(7:(ItemNo+6)), cols=c(1:4, 6:9),gridExpand =T, stack = FALSE)
    addStyle(wb, "ItemAnalyses", border_general, rows=c(5:(7+n_levels)), cols=c(11:14),gridExpand =T, stack = FALSE)
    addStyle(wb, "ItemAnalyses", border_head, rows=c(5,6), cols=c(1:4, 6:9), gridExpand =T, stack = FALSE)

    # "Scores"
    if ("Scores" %in% names(wb)==F) addWorksheet(wb,"Scores")
    writeFormula(wb, sheet="Scores", x = paste0(Skill,"!A1"), startCol = "A", startRow = 1)     # write TestName in Cell A1 from "reading" sheet into "Scores" sheet
    writeFormula(wb, sheet="Scores", x = paste0(Skill,"!A2"), startCol = "A", startRow = 2)
    writeFormula(wb, sheet="Scores", x = paste0(Skill,"!A4"), startCol = "A", startRow = 3)     # write administration date(s) in Cell A4 from "reading" sheet into "Scores" sheet
    writeData(wb, "Scores", dSorted[,1:5], startCol = 1, startRow = 7, rowNames = F, colNames=T)
    writeData(wb, "Scores", cbind(dSorted[,(n_demo+1+ItemNo):(n_demo+3+ItemNo)],rep(0,PerNo),dSorted[,(ncol(dSorted)-1):ncol(dSorted)]),
              startCol = ifelse(Skill=="Reading","J","P") ,startRow = 7)
    # create the bounds table
    Level<-c(paste0("Sub-",ReportedLevels[1]),ReportedLevels)
    UpperBound_R<-round(seq(20,ItemNo,length=n_levels+1),digits = 0)
    LowerBound_R<-c(0,UpperBound_R[1:n_levels]+1)
    bounds_R<-data.frame(Level, LowerBound_R, UpperBound_R)
    writeData(wb, "Scores", bounds_R, startCol = 1, startRow = 10+PerNo, rowNames = F, colNames = T)


    z<-NULL
    p<-paste0("A",11+PerNo)
    for (k in 1:n_levels)
    {
      z<-paste0(z,"IF(K", 8:(7+PerNo), ">C",11+PerNo+n_levels-k, ", A", 12+PerNo+n_levels-k,",")
      p<-paste0(p,")")
    }
    z1<-paste0(z,p)

    class(z1)<-c(class(z1), "formula")
    writeData(wb, "Scores", x=z1, startCol = 13, startRow = 8)


  #### Listening ####
  Skill<-"Listening"
  # Read in response file
  response<-ReadCSVFile(paste0(MainPath,"/"),paste0(TestName, "_", Skill))
  write.csv(response, paste0(CleanPath,paste0(TestName, "_", Skill,".csv")),row.names = F)
  write.csv(response, paste0(OriginalPath,paste0(TestName, "_", Skill,".csv")),row.names = F)
  ItemNo<-ncol(response)-n_demo
  PerNo<-nrow(response) #number of examinees
  RandomItemIDs<-colnames(response)[-c(1:n_demo)] #unique item ID

  if (PL_file==F)
  {
    #### Read in proficiency level file form ACTesting ####
    OriPLs<-ReadCSVFile(paste0(MainPath,"/"),paste(Language, Skill, ProfVersion_l, "levels", sep=" "))
    write.csv(OriPLs, paste0(CleanPath,paste(Language, Skill, ProfVersion_l, "levels", sep=" "),".csv"),row.names = F)
    PL<-matrix(,ncol=4, nrow=ItemNo)
    word<-ifelse(Skill=="Reading","read_","listen_")
    PL[,1]<-c(paste0(word,"0",1:9),paste0(word,10:ItemNo))
    if (ProfScale=="ACTFL")
    {
      PL[,-1]<-as.character(OriPLs$ACTFL)
    } else PL[,-1]<-as.character(OriPLs$ILR)
    colnames(PL)<-c("ItemNo","Original_Prof_Level","Final_Prof_Level","Reported_Prof_Level")
    ItemPLsFile <- paste(TestName,"_",Skill,"_PLs",sep = "")
    write.csv(PL, paste(CleanPath,ItemPLsFile, ".csv", sep = ""), row.names = F)
  } else {
    ItemPLsFile <- paste(TestName,"_",Skill,"_PLs",sep = "")
    PL<-ReadCSVFile(CleanPath, ItemPLsFile)
    PL<-as.matrix(PL)
  }

  #### Scoring and assign proficiency levels ####
  if (ProfScale == "ACTFL"){
    # allowed values for ACTFL proficiency scale.
    allowedValues <- c("NL", "NM", "NH", "NL/NM", "NM/NH", "Novice", "NH/IL",
                       "IL", "IM", "IH", "IL/IM", "IM/IH", "Intermediate", "IH/AL",
                       "AL", "AM", "AH", "AL/AM", "AM/AH", "Advanced", "S", "D",
                       "DNS")
  } else if(ProfScale == "ILR"){
    ## allowed values for ILR proficiency scale.
    allowedValues <- c("0", "0+", "0/0+", "0+/1", "1", "1+", "1/1+", "1+/2",
                       "2", "2+", "2/2+", "2+/3", "3", "3+", "3/3+", "3+/4",
                       "4", "4+", "4/4+", "4+/5", "5", "DNS")
  } else{
    # Stop the program if proficiency scale is not "ACTFL" or "ILR."
    stop(paste("Invalid Proficiency Scale. Score calculations were stopped."))
  }

  c <- unique(PL[,4]) %in% allowedValues

  # Check if the values in the file with test proficiency levels match the ones in allowedValues variable.
  # Stop if there is a mismatch, continue otherwise.
  if ("FALSE" %in% c)
  {
    stop (paste("Invalid values assigned to proficiency levels") )
  } else{
    ReportedLevels <- unique(PL[,4]) [unique(PL[,4]) != "DNS"]
    ReportedLevels <- ReportedLevels[order(match(ReportedLevels, allowedValues))]
    print(paste("The following are proficiency levels in", Skill, ":",
                paste(ReportedLevels, collapse = ", "), collapse = " "))
  }

  # Replace column names, e.g. Q1 to either read_01 or listen_01, etc.
  colnames(response)[-c(1:n_demo)] <- PL[,1]
  # Remove items which are marked as "DNS" as "Do not score" from calculations.
  if ("DNS" %in% as.character(PL[,4])){
    pl <- PL[PL[,4] == "DNS",1] # name of DNS item
    response_pure <- response[, -c(1:n_demo)] # response matrix without demographic data
    response_pure <- response_pure[, !(names(response_pure) %in% pl)]
    print(paste("Dropped from calculations - item(s):",
                paste(pl, collapse = ", "), collapse = " "))
  } else{

    response_pure <- response[, -c(1:n_demo)]
    print("All items are taken into account in the calculations.")
  }

  ## Count Total Omitted, Total Correct, and Total Correct Percent for
  ## Reading or Listening data file.

  Omitted <- apply(response_pure, 1, FUN = function(x) length(which(is.na(x)))+length(which(x=="NA"))) # -1 is omitted in csv, however NA here
  score<-response_pure
  if (length(keychange_l)!=0)
  {
    for (k in 1:length(keychange_r))
    {
      score[,keychange_l[[k]][1]]<-ifelse(response_pure[,keychange_l[[k]][1]] %in% keychange_l[[k]][-1],1,0)
    }
  }
  score[,which(PL[,4]=="DNS")]<-0
  TotalCor <- apply(score, 1, FUN = function(x) length(which(x == 1)))
  TotalCorPcnt <- apply(score, 1, FUN = function(x) (length(which(x == 1)))/ItemNo*100)
  TotalCorPcnt <-round(TotalCorPcnt,digits=2)

  ## Subset dataframe by proficiency levels and counting Number Correct,
  ## Percent Correct, mastery at 75% and above, and mastery at 67% and above
  ## with each proficiency level.

  results <- data.frame()

  for(pls in ReportedLevels){

    ## Redefine the '[' function by adding "drop = FALSE" to prevent from reducing
    ## a matrix into a vector during calculations. For example, if there is only
    ## one item in a proficiency level, the '[' automatically reduces data into
    ## a vector. With "drop = FALSE" change in the "base" package this reduction
    ## is disabled. However, the function is restored after the calculations are
    ## complete so that the change doesn't interfere with the '[' functions defined
    ## for classes other than 'matrix'.

    '[' <- function(...) base::'['(..., drop = FALSE)

    pl <- PL[PL[,4] == pls,1] # items with ProfLevel=pls
    tmp <- score[, colnames(response_pure) %in% pl] # responses belong to pls
    NumCor <- apply(tmp, 1, FUN = function(x) length(which (x == 1)))
    CorPcnt <- apply(tmp, 1, FUN = function(x) length(which(x == 1))/ncol(tmp)*100)
    CorPcnt <- round(CorPcnt,digits = 2)
    Mastery_75Up <- ifelse(CorPcnt > 74.49, "Yes", "No")
    Mastery_67Up <- ifelse(CorPcnt > 66.49, "Yes", "No")
    results <- rbind(results,
                     data.frame(NumCor, CorPcnt, Mastery_75Up, Mastery_67Up))
  }


  # Add proficiency level column for each iteration and merge that column with the results.
  # Split results dataframe into segments the number of which equals to the number of unique proficiency levels in the test section.
  # Assign proficiency level names to split dataframe level names.
  pl_75 <- ifelse(results$Mastery_75Up == "Yes",
                  as.character(rep(ReportedLevels, each = PerNo)), "")
  pl_67 <- ifelse(results$Mastery_67Up == "Yes",
                  as.character(rep(ReportedLevels, each = PerNo)), "")
  results <- cbind(results, pl_75, pl_67)
  results2 <- split(results, rep(1:n_levels,
                                 each = PerNo))
  names(results2) <- ReportedLevels
  results3 <- do.call(cbind, results2)


  # Count the number of mastery levels achieved within each proficiency level.
  data_75Up <- results3[, grep("Mastery_75|pl_75", colnames(results3))]
  Maintain_75Up <- ifelse((apply(data_75Up, 1, FUN = function(x) length(which(x == "Yes")))) == 0,
                          "None", paste("In level(s): ", apply(data_75Up[, grep("pl_75", colnames(data_75Up))],
                                                               1, FUN = function(x) paste(x, collapse = ", "))))

  data_67Up <- results3[, grep("Mastery_67|pl_67", colnames(results3))]
  Maintain_67Up <- ifelse((apply(data_67Up, 1, FUN = function(x) length(which(x == "Yes")))) == 0,
                          "None", paste("In level(s): ", apply(data_67Up[, grep("pl_67", colnames(data_67Up))],
                                                               1, FUN = function(x) paste(x, collapse = ", "))))

  results4 <- results3[, -c(grep("pl_", colnames(results3)))]


  ## Create a dataframe.
  ## Order by examinee's last name, first name in alphabetical order, and then by examinee ID from smallest to largest.


  colnames(response)[-c(1:6)]<-RandomItemIDs  # Return the place holder: RandomItemIDs to the data file variable names


  d <- cbind(response, Omitted, TotalCor, TotalCorPcnt, results4,
             Maintain_75Up, Maintain_67Up)
  d$Maintain_75Up <- gsub("[[:space:]],", "", d$Maintain_75Up)
  d$Maintain_75Up <- gsub(",[[:space:]]$", "", d$Maintain_75Up)
  d$Maintain_67Up <- gsub("[[:space:]],", "", d$Maintain_67Up)
  d$Maintain_67Up <- gsub(",[[:space:]]$", "", d$Maintain_67Up)

  dSorted <- d[with(d, order(d$lastName, d$firstName,d$examineeID)),]

  ## write to Scores folder
  ScoreFile<-paste(TestName,"_",Skill,"_Mastery",sep = "")
  if (file.exists(ScorePath)==F) dir.create(ScorePath)
  write.csv(dSorted, paste(ScorePath,ScoreFile, ".csv", sep = ""),
            row.names = FALSE)

  ## Remove too many omits and repeated examinees
  Cleaned<-dSorted[!dSorted$Omitted>round(ItemNo/4*3,0),]     # removing individuals missing or omited more than 3/4 of questions
  write.csv(Cleaned[,1:(n_demo+ItemNo+3)], paste(jMetrikPath,TestName,"_",Skill,"_jMetrik.csv",sep=""), row.names = FALSE)    # write the cleaned and combined file into a csv file with 8 for missing data


  ## The original '[' base function is being restored here.
  '[' <- function(...) base::'['(..., drop = TRUE)



    #### Create Excel File ####

    # "Listening"
    addWorksheet(wb, Skill)
    writeData(wb, Skill, TestName, startCol = 1, startRow = 1, rowNames = F, colNames=F)          # write TestName in Cell A1
    writeData(wb, Skill, paste("Language",Language,sep= ": "), startCol = 1, startRow = 2, rowNames = F, colNames=F)
    writeData(wb, Skill, paste("Skill",Skill,sep= ": "), startCol = 1, startRow = 3, rowNames = F, colNames=F)
    writeData(wb, Skill, paste(AdminDate,sep= ""), startCol = 1, startRow = 4, rowNames = F, colNames=F)         # write AdminDate(s) in Cell A4

    dSorted_listening<-dSorted
    word1<-ifelse(Skill=="Reading","R","L")
    colnames(dSorted_listening)<-c(colnames(dSorted)[1:n_demo],c(paste0(word1,"0",1:9),paste0(word1,10:ItemNo)),
                                 "Omitted Items", "#Correct", "%Correct",
                                 rep(c("#Correct", "%Correct", "75%&Up", "67%&Up"), n_levels),
                                 "Maintaining 75% and Above Correct", "Maintaining 67% and Above Correct")
    mergeCells(wb, Skill, cols=c(n_demo-1,n_demo), rows=6) # Proficiency levels
    mergeCells(wb, Skill, cols=c(n_demo+ItemNo+2,n_demo+ItemNo+3), rows=6) # Total 44 items
    head_sub<-rep(0,n_levels*4)
    for (k in 1:n_levels)
    {
      head_sub[(4*(k-1)+1):(4*k)]<-c(paste0(ReportedLevels[k]," (",length(which(PL[,4]==ReportedLevels[k]))," items)"),"","","")
      mergeCells(wb, Skill, cols=c(n_demo+ItemNo+4+(k-1)*4,n_demo+ItemNo+4*k+3),rows=6)
    }
    mergeCells(wb, Skill, cols=c(n_demo+ItemNo+4+4*n_levels, n_demo+ItemNo+5+4*n_levels), rows=6)
    head<-data.frame(t(c("Proficiency Levels", "", PL[,4], "", paste0("Total (", ItemNo, " items)"), "",
                         head_sub, paste0("Listening Comprehension (", ItemNo, " items)"))))
    writeData(wb, Skill, head, startCol = n_demo-1, startRow = 6, rowNames = F, colNames=F)
    writeData(wb, Skill, dSorted_listening, startCol = 1, startRow = 7, rowNames = F, colNames=T,keepNA = T)

    addStyle(wb, Skill, border_general, rows=8:(PerNo+7), cols=c(1:(n_demo+ItemNo+5+4*n_levels)), gridExpand =T, stack = FALSE)
    addStyle(wb, Skill, border_head, rows=c(6,7), cols=c(1:(n_demo+ItemNo+5+4*n_levels)), gridExpand = T, stack = FALSE)
    setColWidths(wb, Skill, cols=c(2:(n_demo+ItemNo+5+4*n_levels)), widths = "auto", ignoreMergedCells = T)

    # "SS_Listening"
    addWorksheet(wb,"SS_Listening")
    # index for standard setting
    ind_ss<-sort.int(dSorted$TotalCor,index.return = T)$ix
    hide<-c(grep("75Up",colnames(dSorted)),grep("67Up",colnames(dSorted)))
    ss<-dSorted[ind_ss,-hide]
    writeData(wb,"SS_Listening",ss, startCol = 1, startRow = 2, rowNames = F, colNames=T,keepNA = T)

    mergeCells(wb, "SS_Listening", cols=c(n_demo+ItemNo+2, n_demo+ItemNo+3), rows=1) # merge for "Total 44 items"
    for (k in 1:n_levels)
    {
      head_sub[(2*k-1):(2*k)]<-c(paste0(ReportedLevels[k]," (",length(which(PL[,4]==ReportedLevels[k]))," items)"),"")
      mergeCells(wb, "SS_Listening", cols=c(n_demo+ItemNo+4+(k-1)*2,n_demo+ItemNo+2*k+3),rows=1)
    }
    head<-data.frame(t(c(paste0("Total (", ItemNo, " items)"), "",head_sub)))
    writeData(wb, "SS_Listening", head, startCol= n_demo+ItemNo+2, startRow=1, rowNames = F, colNames=F)

    addStyle(wb, "SS_Listening", border_general, rows=3:(PerNo+2), cols=c(1:(n_demo+ItemNo+3+2*n_levels)), gridExpand =T, stack = FALSE)
    addStyle(wb, "SS_Listening", border_head, rows=c(1,2), cols=c(1:(n_demo+ItemNo+3+2*n_levels)), gridExpand = T, stack = FALSE)
    conditionalFormatting(wb, "SS_Listening", cols=c((n_demo+ItemNo+4):(n_demo+ItemNo+3+2*n_levels)), rows=3:(PerNo+2),
                      rule=">66", style=bold)


    # "ItemAnalyses"
    pl_xlsx<-cbind(PL[,c(1,2,4)],NA)
    colnames(pl_xlsx)<-c("Item No.", "Original PL", "Final PL", "Comments")
    writeData(wb, sheet="ItemAnalyses", pl_xlsx, startCol = ifelse(Skill=="Reading", "A", "F"), startRow = 6)

    levelcount<-NULL
    for (i in 1:n_levels)
    {
      levelcount[i]<-length(which(PL[,4]==ReportedLevels[i]))
    }
    levelsum<-data.frame(c(ReportedLevels,"Total"),c(levelcount,ItemNo))
    colnames(levelsum)<-c("Prof.Lvl","# of Items")
    writeData(wb, sheet="ItemAnalyses", Skill, startCol= ifelse(Skill=="Reading", "K", "M"), startRow = 5)
    writeData(wb, sheet="ItemAnalyses", levelsum, startCol = ifelse(Skill=="Reading", "K", "M"), startRow = 6)

    # "Scores"
    writeData(wb, "Scores", cbind(dSorted[,(n_demo+1+ItemNo):(n_demo+3+ItemNo)],rep(0,PerNo),dSorted[,(ncol(dSorted)-1):ncol(dSorted)]),
              startCol = ifelse(Skill=="Reading","J","P") ,startRow = 7)
    # create the bounds table
    Level<-c(paste0("Sub-",ReportedLevels[1]),ReportedLevels)
    UpperBound_L<-round(seq(20,ItemNo,length=n_levels+1),digits = 0)
    LowerBound_L<-c(0,UpperBound_L[1:n_levels]+1)
    bounds_L<-cbind(Level, LowerBound_R, UpperBound_R)
    writeData(wb, "Scores", bounds_R, startCol = 6, startRow = 10+PerNo, rowNames = F, colNames = T)

    # function to get levels after standard setting
    z<-NULL
    p<-paste0("F", 11+PerNo)
    for (k in 1:n_levels)
    {
      z<-paste0(z,"IF(Q", 8:(7+PerNo), ">H",11+PerNo+n_levels-k, ", F", 12+PerNo+n_levels-k,",")
      p<-paste0(p,")")
    }
    z1<-paste0(z,p)

    class(z1)<-c(class(z1), "formula")
    writeData(wb, "Scores", x=z1, startCol = 19, startRow = 8)

    mergeCells(wb, "Scores", rows =6, cols= 6:9)
    writeData(wb, "Scores", "AAPPL Scores", startCol = "F", startRow = 6, colNames = F)
    mergeCells(wb, "Scores", cols=10:15, rows = 6)
    writeData(wb, "Scores", paste0("Reading Comprehension (", ItemNo, " items)"), startCol = "J", startRow = 6)
    mergeCells(wb, "Scores", cols=16:21, rows = 6)
    writeData(wb, "Scores", paste0("Listening Comprehension (", ItemNo, " items)"), startCol = "P", startRow = 6)

    # Speaking and Writing
    writeData(wb, "Scores", t(c("ILS-Speaking", "Converted Speaking", "PW-Writing", "Converted Writing")), startRow = 7, startCol = "F",colNames = F)
    z_s<-paste0('IF(F', 8:(7+PerNo), '= "A", "IH*", IF(F', 8:(7+PerNo), '="I-5", "IH", IF(F',
                8:(7+PerNo), '= "I-4", "IM", IF(F', 8:(7+PerNo), '="I-3", "IM", IF(F',
                8:(7+PerNo), '= "I-2", "IM", IF(F', 8:(7+PerNo), '="I-1", "IL", IF(F',
                8:(7+PerNo), '= "N-4", "NH", IF(F', 8:(7+PerNo), '="Below N4", "Sub-NH", "n/a"')
    class(z_s)<-c(class(z_s),"formula")
    writeData(wb, "Scores", x=z_s, startCol = "G", startRow = 8)
    z_w<-paste0('IF(H', 8:(7+PerNo), '= "A", "IH*", IF(H', 8:(7+PerNo), '="I-5", "IH", IF(F',
                8:(7+PerNo), '= "I-4", "IM", IF(H', 8:(7+PerNo), '="I-3", "IM", IF(H',
                8:(7+PerNo), '= "I-2", "IM", IF(H', 8:(7+PerNo), '="I-1", "IL", IF(H',
                8:(7+PerNo), '= "N-4", "NH", IF(H', 8:(7+PerNo), '="Below N4", "Sub-NH", "n/a"')
    class(z_w)<-c(class(z_w),"formula")
    writeData(wb, "Scores", x=z_w, startCol = "I", startRow = 8)


    addStyle(wb, "Scores", border_general, cols=c(1:21), rows = c(8:(7+PerNo)),gridExpand =T, stack = FALSE)
    addStyle(wb, "Scores", border_head, cols=c(1:21), rows=c(6,7), gridExpand =T, stack = FALSE)
    setColWidths(wb, "Scores", cols=c(1:21), widths = "auto", ignoreMergedCells = T)


    # "ReportedScores"
    addWorksheet(wb, "ReportedScores")
    writeData(wb, "ReportedScores", paste(TestName,sep= ""), startCol = 1, startRow = 1, rowNames = F, colNames=F)          # write TestName in Cell A1
    writeData(wb, "ReportedScores", paste("Language",Language,sep= ":"), startCol = 1, startRow = 2, rowNames = F, colNames=F)
    writeData(wb, "ReportedScores", paste(AdminDate,sep= ""), startCol = 1, startRow = 3, rowNames = F, colNames=F)         # write AdminDate(s) in Cell A3
    writeData(wb, "ReportedScores", dSorted[,1:3], startCol = 1, startRow = 7, rowNames = F, colNames=T)

    s<-paste0("Scores!G",8:(7+PerNo))
    class(s)<-c(class(s),"formula")
    writeData(wb, "ReportedScores", x=s, startCol = "D", startRow = 8)
    w<-paste0("Scores!I",8:(7+PerNo))
    class(w)<-c(class(w),"formula")
    writeData(wb, "ReportedScores", x=w, startCol = "E", startRow = 8)
    r<-paste0("Scores!M",8:(7+PerNo))
    class(r)<-c(class(r),"formula")
    writeData(wb, "ReportedScores", x=r, startCol = "F", startRow = 8)
    l<-paste0("Scores!S",8:(7+PerNo))
    class(l)<-c(class(l),"formula")
    writeData(wb, "ReportedScores", x=l, startCol = "G", startRow = 8)


    novice<-paste0('COUNTIF(D', 8:(7+PerNo), ':G', 8:(7+PerNo), ',"NH")+COUNTIF(D', 8:(7+PerNo), ':G', 8:(7+PerNo), ',"Sub-NH")')
    intermediate<-paste0('COUNTIF(D', 8:(7+PerNo), ':G', 8:(7+PerNo), ',"IL")+COUNTIF(D', 8:(7+PerNo), ':G', 8:(7+PerNo), ',"IM")')
    advanced<-paste0('COUNTIF(D', 8:(7+PerNo), ':G', 8:(7+PerNo), ',"IH")+COUNTIF(D', 8:(7+PerNo), ':G', 8:(7+PerNo), ',"IH*")')
    Sum<-paste0('SUM(I', 8:(7+PerNo), ":K", 8:(7+PerNo))
    profile<-paste0('IF(L', 8:(7+PerNo), '<4, "NS", IF(I', 8:(7+PerNo), '=4, 1, IF(I',
                    8:(7+PerNo), '=3, 2, IF(I', 8:(7+PerNo), '=2, 1, IF(K', 8:(7+PerNo), '>1, 5, 4')
    class(novice)<-c(class(novice), "formula")
    class(intermediate)<-c(class(intermediate), "formula")
    class(advanced)<-c(class(advanced), "formula")
    class(Sum)<-c(class(Sum), "formula")
    class(profile)<-c(class(profile), "formula")
    writeData(wb, "ReportedScores", x=novice, startCol = "I", startRow = 8, colNames = F)
    writeData(wb, "ReportedScores", x=intermediate, startCol = "J", startRow = 8, colNames = F)
    writeData(wb, "ReportedScores", x=advanced, startCol = "K", startRow = 8, colNames = F)
    writeData(wb, "ReportedScores", x=Sum, startCol = "L", startRow = 8, colNames = F)
    writeData(wb, "ReportedScores", x=profile, startCol = "M", startRow = 8, colNames = F)

    addStyle(wb, "ReportedScores", border_general, cols=c(1:13), rows = c(8:(7+PerNo)),gridExpand =T, stack = FALSE)
    addStyle(wb, "ReportedScores", border_head, cols=c(1:13), rows=c(6,7), gridExpand =T, stack = FALSE)

    # write head
    writeData(wb, "ReportedScores", x=t(c("Speaking", "Writing", "Reading", "Listening", "",
                                           "Novice", "Intermediate", "Advanced", "Sumn", "Profile Score")),
              startCol = "D", startRow = 7, colNames = F)
    mergeCells(wb, "ReportedScores", rows =6, cols= 4:5)
    mergeCells(wb, "ReportedScores", rows =6, cols= 6:7)
    mergeCells(wb, "ReportedScores", rows =6, cols= 9:11)
    writeData(wb, "ReportedScores", t(c("AAPPL Scores", "","AC Scores", "","", "Number of Skill Subscores")),
              startCol = "D", startRow = 6, colNames = F)


    # "Profile Score by Group"
    addWorksheet(wb,"Profile Score by Group")
    for (j in c(1:7,9:13))
    {
      v<-paste0("ReportedScores!", LETTERS[j], 7:(7+PerNo))
      class(v)<-c(class(v), "formula")
      writeData(wb, "Profile Score by Group", x=v, startCol = j, startRow = 7, rowNames = F, colNames=F)
    }
    writeData(wb, "Profile Score by Group", dSorted[,6], startCol = "N", startRow = 8, colNames = T)
    writeData(wb, "Profile Score by Group", paste(TestName,sep= ""), startCol = 1, startRow = 1, rowNames = F, colNames=F)          # write TestName in Cell A1
    writeData(wb, "Profile Score by Group", paste("Language",Language,sep= ":"), startCol = 1, startRow = 2, rowNames = F, colNames=F)
    writeData(wb, "Profile Score by Group", paste(AdminDate,sep= ""), startCol = 1, startRow = 3, rowNames = F, colNames=F)
    writeData(wb, "Profile Score by Group", "Inheritage", startCol = "N", startRow = 7, rowNames = F, colNames=F)

    mergeCells(wb, "Profile Score by Group", rows =6, cols= 4:5)
    mergeCells(wb, "Profile Score by Group", rows =6, cols= 6:7)
    mergeCells(wb, "Profile Score by Group", rows =6, cols= 9:11)
    writeData(wb, "Profile Score by Group", t(c("AAPPL Scores", "","AC Scores", "","", "Number of Skill Subscores")),
              startCol = "D", startRow = 6, colNames = F)

    addStyle(wb, "Profile Score by Group", border_general, cols=c(1:14), rows = c(8:(7+PerNo)),gridExpand =T, stack = FALSE)
    addStyle(wb, "Profile Score by Group", border_head, cols=c(1:14), rows=c(6,7), gridExpand =T, stack = FALSE)

    worksheetOrder(wb)<-c(1,6,3,4,5,8,9,2,7)
    saveWorkbook(wb, file = paste0(CleanPath,"Scored",TestName,".xlsx"), overwrite = TRUE)

}
