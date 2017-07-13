#' @name Tucker_equating
#' @title A function to do the Tucker linear equating.
#' @description It produces ItemStats file (for both old test and new test) and equating table.
#' Plots of liking item score and total score distribution are also developed.
#' @param MainPath The folder where the response files are stored in
#' @param Xpath Path of the old response file is stored in
#' @param Ypath Path of the new response file is stored in
#' @param XName Name of the X test (without "_Listening" or "_Reading")
#' @param XName Name of the Y test
#' @param n_demoX Number of demographic variables in X. Default value is 6.
#' @param n_demoY Number of demographic variables in Y.
#' @param DNS_X_r Vector of not score items in X reading test. Default is NULL.
#' @param DNS_X_l Vector of not score items in X listening test.
#' @param DNS_Y_r Vector of not score items in Y reading test.
#' @param DNS_Y_l Vector of not score items in Y listening test.
#' @param keychange_X_r List of key change for X reading test. Default is NULL.
#' The first element is the item index, the rest are the keys.
#' @param keychange_X_l List of key change for X listening test.
#' @param keychange_Y_r List of key change for Y reading test.
#' @param keychange_Y_l List of key change for Y listening test.
#' @param equate_coef List of Tucker linear equating coefficience. First element is intercept, second is slope.
#' e.g. c_2015 equates 2016 test score to 2015. When using this function, just fill in proper equating coefficience.
#' @return Stats for both forms of test and the equating table.
#' @export
#' @examples
#' Tucker_equating(MainPath, XPath, YPath, XName, YName, n_demoX=6, n_demoY=6,
#'  DNS_X_r=NULL, DNS_X_l=c(1,5), DNS_Y_r=NULL, DNS_Y_l=NULL,
#'  keychange_X_r=list(c(2,1,4),c(13, 2)), keychange_X_l=NULL, keychange_Y_r=NULL, kechange_Y_l=NULL,
#'  equate_coef=list(c_2015=c(-2.282018,1.052596),
#'                 c_2014=c(-1.848358,1.014028),
#'                 c_2013=c(2.308507,0.9523797),
#'                 c_2012=c(3.089623,0.9053153)))


Tucker_equating<-function(MainPath, XPath, YPath, XName, YName, n_demoX=6, n_demoY=6,
                          DNS_X_r=NULL, DNS_X_l=NULL, DNS_Y_r=NULL, DNS_Y_l=NULL,
                          keychange_X_r=NULL, keychange_X_l=NULL, keychange_Y_r=NULL, keychange_Y_l=NULL,
                          equate_coef=list(c_2015=c(-2.282018,1.052596),
                                           c_2014=c(-1.848358,1.014028),
                                           c_2013=c(2.308507,0.9523797),
                                           c_2012=c(3.089623,0.9053153)))
{
 library(equate)
 library(psych)
 for (Skill in c("Reading", "Listening"))
 {
   # create path for equating
   EquatingPath <- paste0(MainPath,"/Equating/")
   if (file.exists(EquatingPath)==F)
   {
     dir.create(file.path(EquatingPath))
   }
   #### compute Item Stats ####
   # Read in response matrices
   X_Data <- ReadCSVFile(XPath, paste0(XName, "_", Skill))
   Y_Data <- ReadCSVFile(YPath, paste0(YName, "_", Skill))

   # Ommited response
   X_Responses<-X_Data[,-(1:n_demoX)]
   X_Responses<-replace(X_Responses, X_Responses==-1, 8)
   X_Responses<-replace(X_Responses, X_Responses=="NA", 8)
   X_Responses<-replace(X_Responses, is.na(X_Responses), 8)
   Y_Responses<-Y_Data[,-(1:n_demoY)]
   Y_Responses<-replace(Y_Responses, Y_Responses==-1, 8)
   Y_Responses<-replace(Y_Responses, Y_Responses=="NA", 8)
   Y_Responses<-replace(Y_Responses, is.na(Y_Responses), 8)

   # keychange
   if (Skill=="Reading")
   {
     keychange_X<-keychange_X_r
   } else keychange_X<-keychange_X_l

   if (length(keychange_X)!=0)
   {
     for (k in 1:length(keychange_X))
     {
       ind<-which(X_Responses[,keychange_X[[k]][1]] %in% keychange_X[[k]][-1])
       X_Responses[ind,keychange_X[[k]][1]]<-1
       X_Responses[-ind,keychange_X[[k]][1]]<-0
     }
   }

   if (Skill=="Reading")
   {
     keychange_Y<-keychange_Y_r
   } else keychange_Y<-keychange_Y_l

   if (length(keychange_Y)!=0)
   {
     for (k in 1:length(keychange_Y))
     {
       ind<-which(Y_Responses[,keychange_Y[[k]][1]] %in% keychange_Y[[k]][-1])
       Y_Responses[ind,keychange_Y[[k]][1]]<-1
       Y_Responses[-ind,keychange_Y[[k]][1]]<-0
     }
   }

   # DNS
   # Build matrix of attempted responses and scored responses for Form X
   if(Skill=="Reading")
   {
     DNS_X<-DNS_X_r
   } else DNS_X<-DNS_X_l

   X_Responses[,DNS_X]<- -1
   ind_DNS<-which(X_Responses==-1)
   if (length(ind_DNS)<1) ind_DNS=NULL
   ind_missing<-which(X_Responses==8)
   if (length(ind_missing)<1) ind_missing=NULL

   ind_wrong<-which(X_Responses!=1)

   X_ItemAttempts<-as.matrix(X_Responses)
   X_ItemAttempts[c(ind_DNS,ind_missing)]<- 0
   X_ItemAttempts[ind_wrong]<- 1

   X_ItemScores<-X_ItemAttempts
   X_ItemScores[ind_wrong]<-0

   N_x      <- nrow(X_Responses) # sample size
   Nitems_x <- ncol(X_Responses) # test length

   # Compute Percent "Attempted" Vector for Items in Form X
   X_UnitVec <- rep(1, N_x)
   X_Prop_Attempted <- (1/N_x)*(X_UnitVec %*% X_ItemAttempts) # 0 for DNS

   # Compute Percent Correct (Item Difficulty) for Form X
   X_Prop_Correct <- (1/N_x)*(X_UnitVec %*% X_ItemScores)   # 0 for DNS

   # Build matrix of attempted responses and scored responses for Form Y
   if(Skill=="Reading")
   {
     DNS_Y<-DNS_Y_r
   } else DNS_Y<-DNS_Y_l

   Y_Responses[,DNS_Y]<- -1
   ind_DNS<-which(Y_Responses==-1)
   if (length(ind_DNS)<1) ind_DNS=NULL
   ind_missing<-which(Y_Responses==8)
   if (length(ind_missing)<1) ind_missing=NULL

   ind_wrong<-which(Y_Responses!=1)

   Y_ItemAttempts<-as.matrix(Y_Responses)
   Y_ItemAttempts[c(ind_DNS,ind_missing)]<- 0
   Y_ItemAttempts[ind_wrong]<- 1

   Y_ItemScores<-Y_ItemAttempts
   Y_ItemScores[ind_wrong]<-0

   N_y      <- nrow(Y_Responses) # sample size
   Nitems_y <- ncol(Y_Responses) # test length

   # Compute Percent "Attempted" Vector for Items in Form Y
   Y_UnitVec <- rep(1, N_y)
   Y_Prop_Attempted <- (1/N_y)*(Y_UnitVec %*% Y_ItemAttempts) # 0 entries for DNS

   # Compute Percent Correct (Item Difficulty) for Form Y
   Y_Prop_Correct <- (1/N_y)*(Y_UnitVec %*% Y_ItemScores)   # 0 for DNS

   #### Linking Item ID ####
   if (length(DNS_X)!=0)
   {
     ItemID_X<-as.numeric(colnames(X_Responses))[-DNS_X]
   } else ItemID_X<-as.numeric(colnames(X_Responses))

   if (length(DNS_Y)!=0)
   {
     ItemID_Y<-as.numeric(colnames(Y_Responses))[-DNS_Y]
   } else ItemID_Y<-as.numeric(colnames(Y_Responses))

   LinkingItemIDs<-NULL
   for (i in 1:length(ItemID_Y))
   {
     try1<-ItemID_Y[i]
     LinkingItemIDs<-c(LinkingItemIDs,ItemID_X[which(ItemID_X==try1)])
   }
   NumLinkingItems<- length(LinkingItemIDs)

   # Construct LInking Items Design Map for Form X
   X_LinkingDesign <- cbind(rep(1,Nitems_x), rep(0,Nitems_x)) # Initialized to [1|0] matrix
   for (j in 1:Nitems_x)
   {
     if (ItemID_X[j] %in% LinkingItemIDs)
     {
       X_LinkingDesign[j,2] <- 1    # Turn on the common items
     }
   }

   Scores_X_XV <- X_ItemScores %*% X_LinkingDesign # Calculate Total Score in V1
   # and Common Item Score in V2

   # Calculate Point-biserial and Biserial Correlations for Form X
   X_Pearson  <- X_Biserial <-rep(0, Nitems_x) # point-biserial & biserial

   for (j in 1:Nitems_x)
   {
     if ( X_Prop_Attempted[j] > 0)
     {   # Calculations only for items not turned off and responded to
       if (sum(X_ItemScores[,j])==N_x | sum(X_ItemScores[,j])==0)
       {
         X_Pearson[j]<-X_Biserial[j]<-0
       } else {
         X_Pearson[j]<-cor(Scores_X_XV[,1]-X_ItemScores[, j],X_ItemScores[, j])
         X_Biserial[j]<-biserial(Scores_X_XV[,1]-X_ItemScores[, j],X_ItemScores[, j])
       }
     }
   }

   # Assemble Table of Descriptive ItemStatistics, Form X
   dim(X_Prop_Attempted)   <- c(Nitems_x, 1)  # Transpose results vector
   dim(X_Prop_Correct)     <- c(Nitems_x, 1)  # ditto
   dim(X_Pearson)          <- c(Nitems_x, 1)  # ditto
   dim(X_Biserial)         <- c(Nitems_x, 1)  # ditto

   X_DescriptiveStatsTable <- cbind(as.numeric(colnames(X_Responses)), X_LinkingDesign[,2], X_Prop_Attempted, X_Prop_Correct,
                                    X_Pearson, X_Biserial)
   DescriptiveStatsColNames <- c("Item ID.", "Linking", "Attempted", "Correct", "R_pbis", "R_bis")
   colnames(X_DescriptiveStatsTable) <- DescriptiveStatsColNames

   X_StatsTablefile <- paste0(EquatingPath,Skill, "_X_ItemStats.csv")
   write.csv(X_DescriptiveStatsTable, X_StatsTablefile, row.names=FALSE)


   # Construct LInking Items Design Map for Form Y
   Y_LinkingDesign <- cbind(rep(1,Nitems_y), rep(0,Nitems_y)) # Initialized to [1|0] matrix
   for (j in 1:Nitems_y)
   {
     if (ItemID_Y[j] %in% LinkingItemIDs)
     {
       Y_LinkingDesign[j,2] <- 1    # Turn on the common items
     }
   }

   Scores_Y_YV <- Y_ItemScores %*% Y_LinkingDesign # Calculate Total Score in V1
   # and Common Item Score in V2

   # Calculate Point-biserial and Biserial Correlations for Form Y
   Y_Pearson  <- Y_Biserial <-rep(0, Nitems_y) # point-biserial & biserial

   for (j in 1:Nitems_y)
   {
     if ( Y_Prop_Attempted[j] > 0)
     {   # Calculations only for items not turned off and responded to
       if (sum(Y_ItemScores[,j])==N_y | sum(Y_ItemScores[,j])==0)
       {
         Y_Pearson[j]<-Y_Biserial[j]<-0
       } else {
         Y_Pearson[j]<-cor(Scores_Y_YV[,1]-Y_ItemScores[, j],Y_ItemScores[, j])
         Y_Biserial[j]<-biserial(Scores_Y_YV[,1]-Y_ItemScores[, j],Y_ItemScores[, j])
       }
     }
   }

   # Assemble Table of Descriptive ItemStatistics, Form Y
   dim(Y_Prop_Attempted)   <- c(Nitems_y, 1)  # Transpose results vector
   dim(Y_Prop_Correct)     <- c(Nitems_y, 1)  # ditto
   dim(Y_Pearson)          <- c(Nitems_y, 1)  # ditto
   dim(Y_Biserial)         <- c(Nitems_y, 1)  # ditto

   Y_DescriptiveStatsTable <- cbind(as.numeric(colnames(Y_Responses)), Y_LinkingDesign[,2], Y_Prop_Attempted, Y_Prop_Correct,
                                    Y_Pearson, Y_Biserial)
   colnames(Y_DescriptiveStatsTable) <- DescriptiveStatsColNames

   Y_StatsTablefile <- paste0(EquatingPath,Skill, "_Y_ItemStats.csv")
   write.csv(Y_DescriptiveStatsTable, Y_StatsTablefile, row.names=FALSE)

   #### Equating ####
   # Freqs of X and V, Pop 1
   neat.x <- freqtab(Scores_X_XV, scales=list(0:Nitems_x,0:NumLinkingItems) )
   # Freqs of y and V, Pop 2
   neat.y <- freqtab(Scores_Y_YV, scales=list(0:Nitems_y,0:NumLinkingItems) )

   jpeg(paste0(EquatingPath,Skill, "_neat_X.jpg"))
   plot(neat.x, xlab="Total Test", ylab="Anchor Test")  # Plot of V against X in Population 1
   dev.off()
   jpeg(paste0(EquatingPath,Skill, "_neat_Y.jpg"))
   plot(neat.y, xlab="Total Test", ylab="Anchor Test")  # Plot of V against Y in Population 2
   dev.off()

   # Tucker Raw Score Equating
   neat.Tucker <- equate(neat.y, neat.x, type='linear', method = 'tucker')

   # Equating Constants

   cat("\nMapping Y to X by Tucker Equating:\n")

   Tucker_slope <- neat.Tucker$coeff[[2]] # Reciprocal of Albano's Slope
   Tucker_icpt  <- neat.Tucker$coeff[[1]]

   cat("Intercept: ", Tucker_icpt,"\n")
   cat("    Slope: ", Tucker_slope,"\n")

   # Equating table
   equating_table<-matrix(, ncol=3+length(equate_coef), nrow=Nitems_y-length(DNS_Y)+1)
   equating_table[,1]<-seq(from = 0, to = Nitems_y-length(DNS_Y), by = 1) # Form Y score
   equating_table[,2]<-Tucker_icpt + Tucker_slope*equating_table[,1] # Form X score
   for (k in 1:length(equate_coef))
   {
     equating_table[,2+k]<-equate_coef[[k]][1] + equate_coef[[k]][2]*equating_table[,1+k]
   }
   equating_table[,3+length(equate_coef)]<- round(equating_table[,2+length(equate_coef)])
   equating_table[,3+length(equate_coef)] <- equating_table[,3+length(equate_coef)] * (equating_table[,3+length(equate_coef)] >= 0.0)

   table_head_sub<-NULL
   name_sub<-"2012"
   for (k in 1:(1+length(equate_coef)))
   {
     table_head_sub[k]<-paste0("Tucker_Equated_", 2011+k, "_Float")
     name_sub<-paste(2012+k,name_sub, sep="_")
   }
   table_head<-c("Form_Y_Score", table_head_sub[(1+length(equate_coef)):1], "Tucker_Equated_2012_Score")
   colnames(equating_table)<-table_head

   write.csv(equating_table,file=paste0(EquatingPath, "Tucker", name_sub,"_",Skill, "EquatingTable.csv"), row.names=FALSE)
 }
}
