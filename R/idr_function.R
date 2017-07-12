#' @name IDR
#' @title A function to prepare for IDR.
#' @description Generate keycheck figures and DIF figures, perform DIF analysis, CTT analysis, and IRT analysis.
#' All the plots and csv files are stored in the folders. The results stored in txt files in jMetrik are stored as outputs of this function.
#' Reading is always analyzed before Listening.
#' @param MainPath The folder where the response files are stored in
#' @param TestName Name of the test
#' @param n_demo number of demographic variables, usually it's 6.
#' @param DIF whether we are going to do DIF detection. The rule of thumb is each group should have more than 50 participants.
#' Thus, the default is F.
#' @return Descriptive summary, dif results, CTT results and IRT results
#' @examples
#' IDR(MainPath, TestName, n_demo=6)

IDR<-function(MainPath, TestName, n_demo=6, DIF=F)
{
  library("psych")
  library("KernSmooth")
  library("difR")
  library("eRm")
  result<-list(NULL,NULL)
  for (Skill in c("Reading","Listening"))
  {
    Data<-ReadCSVFile(paste0(MainPath,"/Data/jMetrik/"),paste0(TestName, "_", Skill,"_jMetrik"))
    ItemNo<-ncol(Data)-3-n_demo
    PerNo<-nrow(Data)
    response<-Data[,(n_demo+1):(n_demo+ItemNo)] # delete columns with demographic information and total score etc.
    name<-colnames(response)
    #### Descriptive Statistics ####
    ctt_score<-ctt_Score(response, key=1)
    sum_score<-ctt_score$score
    score_matrix<-ctt_score$scored
    dist_1<-ctt_Score(response, key=2)$scored
    dist_2<-ctt_Score(response, key=3)$scored
    dist_3<-ctt_Score(response, key=4)$scored

    sum_describe1<- matrix(round(describe(sum_score),digits = 4),ncol=1)
    sum_describe1[13]<-round(sd(sum_score),digits=4)
    name_describe<-c("N","Min","Max","Mean","St.Dev.", "Skewness", "Kurtosis")
    sum_describe<-data.frame(cbind(name_describe,sum_describe1[c(2,8,9,3,13,11,12)]))
    colnames(sum_describe)<-c("Statistic", "Value")
    rownames(sum_describe)<-NULL

    #### Kernel Desnsity ###
    den<-density(sum_score,kernel = "gaussian",bw="bcv")
    bw<-den$bw
    #### Key Check ####
    # create folder to store key check plots
    KeycheckPath<-paste0(MainPath, "/UploadFiles/", Skill, "/keycheck/")
    if (file.exists(KeycheckPath)==F)
    {
      if (file.exists(paste0(MainPath,"/UploadFiles/"))==F)
      {
        dir.create(paste0(MainPath,"/UploadFiles/"))
        dir.create(paste0(MainPath,"/UploadFiles/",Skill,"/"))
        dir.create(KeycheckPath)
      } else {
        if (file.exists(paste0(MainPath,"/UploadFiles/",Skill,"/"))==F)
        {
          dir.create(paste0(MainPath,"/UploadFiles/",Skill,"/"))
          dir.create(KeycheckPath)
        } else dir.create(KeycheckPath)
      }
    }


    blue <- rgb(30,144,255,alpha=60,maxColorValue=255)
    red <- rgb(255, 0, 30, alpha=60, maxColorValue = 255)
    grey <- rgb(50,0,50, alpha=60, maxColorValue = 255)
    green<- rgb(0,200,0, maxColorValue = 255)
    exp_score<-matrix(,nrow=ItemNo,ncol=51) # expected score
    for (j in 1:ItemNo)
    {
      jpeg(paste0(KeycheckPath,colnames(response)[j],".jpg"))
      fit<-locpoly(sum_score, score_matrix[,j], drv = 0L, degree=1, kernel = "normal",
                   bandwidth=bw, gridsize = 51L, bwdisc = 51,
                   range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
      exp_score[j,]<-fit$y
      plot(fit$x,fit$y,type="l", ylim=c(0,1), #main=paste0("ICC for ",colnames(response)[j]),
           xlim=range(sum_score), xlab="Total Correct",ylab="Probability",lwd=2)

      fit<-locpoly(sum_score, dist_1[,j], drv = 0L, degree=1, kernel = "normal",
                   bandwidth=bw, gridsize = 51L, bwdisc = 51,
                   range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
      lines(fit$x,fit$y,ylim=c(0,1),xlim=range(sum_score),col="red",lty=2,lwd=2)

      fit<-locpoly(sum_score, dist_2[,j], drv = 0L, degree=1, kernel = "normal",
                   bandwidth=bw, gridsize = 51L, bwdisc = 51,
                   range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
      lines(fit$x,fit$y,ylim=c(0,1),xlim=range(sum_score),col="blue",lty=3,lwd=2)

      fit<-locpoly(sum_score, dist_3[,j], drv = 0L, degree=1, kernel = "normal",
                   bandwidth=bw, gridsize = 51L, bwdisc = 51,
                   range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
      lines(fit$x,fit$y,ylim=c(0,1),xlim=range(sum_score),col=green,lty=4,lwd=2)

      #legend("topleft", legend=c("Key", "Distractor 1", "Distractor 2", "Distractor 3"),
      #       lty=c(1,2,2,2), col=c("black", "red", "blue", green),bty="n")
      polygon(x=c(den$x,rev(den$x)),y=c(den$y*10,rep(0,length(den$x))),col=blue,border=NA)
      dev.off()
    }

    #### Test Characteristic Curve ####
    # plot(fit$x,colSums(exp_score),type="l", main="Test Characteristic Curve",xlim=range(sum_score),
    #      xlab="Total Correct",ylab="True Score")
    if (DIF)
    {
      #### DIF Detection Figures ####
      #### gender ####
      ind_m<-which(Data$gender=="M") # index for male
      sum_score_m<-sum_score[ind_m]
      sum_score_f<-sum_score[-ind_m]
      score_matrix_m<-score_matrix[ind_m,]
      score_matrix_f<-score_matrix[-ind_m,]

      DifPath_g1<-paste0(MainPath, "/UploadFiles/", Skill, "/dif/gender/") # in upload file
      if (file.exists(DifPath_g1)==F)
      {
        if (file.exists(paste0(MainPath, "/UploadFiles/", Skill, "/dif/"))==F)
        {
          dir.create(paste0(MainPath, "/UploadFiles/", Skill, "/dif/"))
          dir.create(DifPath_g1)
        } else dir.create(DifPath_g1)
      }

      den_m<-density(sum_score_m,kernel = "gaussian", bw="bcv")
      den_f<-density(sum_score_f,kernel = "gaussian", bw="bcv")
      for (j in 1:ItemNo)
      {
        jpeg(paste0(DifPath_g1,colnames(response)[j],".jpg"))
        fit<-locpoly(sum_score_f, score_matrix_f[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        plot(fit$x,fit$y,type="l", ylim=c(0,1), #main=paste0("Gender DIF for ",colnames(response)[j]),
             xlim=range(sum_score), xlab="Total Correct",ylab="Probability",lwd=2)

        fit<-locpoly(sum_score_m, score_matrix_m[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        lines(fit$x,fit$y,type="l", ylim=c(0,1),xlim=range(sum_score),lty=2,col="red",lwd=2)

        #legend("topleft", legend=c("Focal Group (F)", "Reference Group (M)"),
        #       lty=c(1,2), col=c("black", "red"),bty="n")
        polygon(x=c(den_m$x,rev(den_m$x)),y=c(den_m$y*length(ind_m)*20/PerNo,rep(0,length(den_m$x))),
                col=red,border=NA, xlim=range(sum_score))
        polygon(x=c(den_f$x,rev(den_f$x)),y=c(den_f$y*(PerNo-length(ind_m))*20/PerNo,rep(0,length(den_f$x))),
                col=grey,border=NA, xlim=range(sum_score))

        dev.off()
      }

      DifPath_g2<-paste0(MainPath, "/", Skill, "/DIF/gender/") # in Reading or Listening file
      if (file.exists(DifPath_g2)==F)
      {
        if (file.exists(paste0(MainPath, "/", Skill, "/DIF/"))==F)
        {
          if (file.exists(paste0(MainPath, "/",Skill))==F)
          {
            dir.create(paste0(MainPath,"/", Skill,"/"))
            dir.create(paste0(MainPath, "/", Skill, "/DIF/"))
            dir.create(DifPath_g2)
          } else {
            dir.create(paste0(MainPath, "/", Skill, "/DIF/"))
            dir.create(DifPath_g2)
          }
        } else dir.create(DifPath_g2)
      }

      for (j in 1:ItemNo)
      {
        jpeg(paste0(DifPath_g2,colnames(response)[j],".jpg"))
        fit<-locpoly(sum_score_f, score_matrix_f[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        plot(fit$x,fit$y,type="l", ylim=c(0,1), #main=paste0("Gender DIF for ",colnames(response)[j]),
             xlim=range(sum_score), xlab="Total Correct",ylab="Probability",lwd=2)

        fit<-locpoly(sum_score_m, score_matrix_m[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        lines(fit$x,fit$y,type="l", ylim=c(0,1),xlim=range(sum_score),lty=2,col="red",lwd=2)

        #legend("topleft", legend=c("Focal Group (F)", "Reference Group (M)"),
        #       lty=c(1,2), col=c("black", "red"),bty="n")

        polygon(x=c(den_m$x,rev(den_m$x)),y=c(den_m$y*length(ind_m)*20/PerNo,rep(0,length(den_m$x))),
                col=red,border=NA, xlim=range(sum_score))
        polygon(x=c(den_f$x,rev(den_f$x)),y=c(den_f$y*(PerNo-length(ind_m))*20/PerNo,rep(0,length(den_f$x))),
                col=grey,border=NA, xlim=range(sum_score))
        dev.off()
      }


      #### heritage_t ####
      ind_y<-which(Data$heritage_t=="Y") # index for heritage students
      sum_score_y<-sum_score[ind_y]
      sum_score_n<-sum_score[-ind_y]
      score_matrix_y<-score_matrix[ind_y,]
      score_matrix_n<-score_matrix[-ind_y,]

      DifPath_h1<-paste0(MainPath, "/UploadFiles/", Skill, "/dif/heritage/")
      if (file.exists(DifPath_h1)==F)
      {
        dir.create(DifPath_h1)
      }

      den_y<-density(sum_score_y,kernel = "gaussian", bw="bcv")
      den_n<-density(sum_score_n,kernel = "gaussian", bw="bcv")
      for (j in 1:ItemNo)
      {
        jpeg(paste0(DifPath_h1,colnames(response)[j],".jpg"))
        fit<-locpoly(sum_score_y, score_matrix_y[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        plot(fit$x,fit$y,type="l", ylim=c(0,1), #main=paste0("Heritage DIF for ", colnames(response)[j]),
             xlim=range(sum_score), xlab="Total Correct",ylab="Probability",lwd=2)

        fit<-locpoly(sum_score_n, score_matrix_n[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        lines(fit$x,fit$y,type="l", ylim=c(0,1),xlim=range(sum_score),lty=2,col="red",lwd=2)

        #legend("topleft", legend=c("Focal Group (Y)", "Reference Group (N)"),
        #       lty=c(1,2), col=c("black", "red"),bty="n")

        polygon(x=c(den_y$x,rev(den_y$x)),y=c(den_y$y*length(ind_y)*20/PerNo,rep(0,length(den_y$x))),
                col=red,border=NA, xlim=range(sum_score))
        polygon(x=c(den_n$x,rev(den_n$x)),y=c(den_n$y*(PerNo-length(ind_y))*20/PerNo,rep(0,length(den_n$x))),
                col=grey,border=NA, xlim=range(sum_score))
        dev.off()
      }

      DifPath_h2<-paste0(MainPath, "/", Skill, "/DIF/heritage_t/") # in Reading or Listening file
      if (file.exists(DifPath_h2)==F)
      {
        if (file.exists(paste0(MainPath, "/", Skill, "/DIF/"))==F)
        {
          if (file.exists(paste0(MainPath, "/", Skill))==F)
          {
            dir.create(paste0(MainPath,"/", Skill,"/"))
            dir.create(paste0(MainPath, "/", Skill, "/DIF/"))
            dir.create(DifPath_h2)
          } else {
            dir.create(paste0(MainPath, "/", Skill, "/DIF/"))
            dir.create(DifPath_h2)
          }
        } else dir.create(DifPath_h2)
      }

      for (j in 1:ItemNo)
      {
        jpeg(paste0(DifPath_h2,colnames(response)[j],".jpg"))
        fit<-locpoly(sum_score_y, score_matrix_y[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        plot(fit$x,fit$y,type="l", ylim=c(0,1), #main=paste0("Heritage DIF for ",colnames(response)[j]),
             xlim=range(sum_score), xlab="Total Correct",ylab="Probability",lwd=2)

        fit<-locpoly(sum_score_n, score_matrix_n[,j], drv = 0L, degree=1, kernel = "normal",
                     bandwidth=bw, gridsize = 51L, bwdisc = 51,
                     range.x<-range(sum_score)+c(-5,5), binned = FALSE, truncate = TRUE)
        lines(fit$x,fit$y,type="l", ylim=c(0,1),xlim=range(sum_score),lty=2,col="red",lwd=2)

        #legend("topleft", legend=c("Focal Group (Y)", "Reference Group (N)"),
        #       lty=c(1,2), col=c("black", "red"),bty="n")

        polygon(x=c(den_y$x,rev(den_y$x)),y=c(den_y$y*length(ind_y)*20/PerNo,rep(0,length(den_y$x))),
                col=red,border=NA, xlim=range(sum_score))
        polygon(x=c(den_n$x,rev(den_n$x)),y=c(den_n$y*(PerNo-length(ind_y))*20/PerNo,rep(0,length(den_n$x))),
                col=grey,border=NA, xlim=range(sum_score))
        dev.off()
      }



      #### DIF Detection table ####

      score_withna<-score_matrix
      score_withna[which(response=="NA")]<-NA
      # DIF for gender
      mh_g<-difMH(Data=score_withna,group=Data$gender,focal.name = "F",correct=F)

      chisq_g<-mh_g$MH
      pvalue_g<-1-pchisq(chisq_g,df=1)
      #n<-
      effectsize_g<--2.35*log(mh_g$alphaMH)
      upper_g<-effectsize_g+1.96*2.35*sqrt(mh_g$varLambda)
      lower_g<-effectsize_g-1.96*2.35*sqrt(mh_g$varLambda)
      etsclass_g<-rep("B",ItemNo)
      for (j in 1:ItemNo)
      {
        if (pvalue_g[j]=="NaN")
        {
          etsclass_g[j]<-"NaN"
        } else {
          if (pvalue_g[j]>0.05 | abs(effectsize_g[j])<1)
          {
            etsclass_g[j]<-"A"
          } else {
            if (pvalue_g[j]<0.05 & abs(effectsize_g[j])>1.5)
            {
              etsclass_g[j]<-"C"
            }
          }
        }
      }
      ind_plus<-which((etsclass_g=="B" | etsclass_g=="C") & effectsize_g>0)
      etsclass_g[ind_plus]<-paste0(etsclass_g[ind_plus],"+")
      dif_full_g<-cbind(name,chisq_g,pvalue_g,effectsize_g,upper_g,lower_g,etsclass_g)
      write.csv(dif_full_g, paste0(DifPath_g2,Skill,"_DIF_gender.csv"),row.names<-F)

      # DIF for heritage
      mh_h<-difMH(Data=score_withna,group=Data$heritage_t,focal.name = "Y",correct=F)
      chisq_h<-mh_h$MH
      pvalue_h<-1-pchisq(chisq_h,df=1)
      #n<-
      effectsize_h<--2.35*log(mh_h$alphaMH)
      upper_h<-effectsize_h+1.96*2.35*sqrt(mh_h$varLambda)
      lower_h<-effectsize_h-1.96*2.35*sqrt(mh_h$varLambda)
      etsclass_h<-rep("B",ItemNo)
      for (j in 1:ItemNo)
      {
        if (pvalue_h[j]=="NaN")
        {
          etsclass_h[j]<-"NaN"
        } else {
          if (pvalue_h[j]>0.05 | abs(effectsize_h[j])<1)
          {
            etsclass_h[j]<-"A"
          } else {
            if (pvalue_h[j]<0.05 & abs(effectsize_h[j])>1.5)
            {
              etsclass_h[j]<-"C"
            }
          }
        }
      }
      ind_plus<-which((etsclass_h=="B" | etsclass_h=="C") & effectsize_h>0)
      etsclass_h[ind_plus]<-paste0(etsclass_h[ind_plus],"+")
      dif_full_h<-cbind(name,chisq_h,pvalue_h,effectsize_h,upper_h,lower_h,etsclass_h)
      write.csv(dif_full_h, paste0(DifPath_h2,Skill,"_DIF_heritage.csv"),row.names=F)

      # create DIF csv file #
      item_id<-rep(name, each=2)
      collection_id<-rep(TestName, 2*ItemNo)
      Type<-rep(c("gender","heritage"),length=2*ItemNo)
      Value<-Group_1_Value<-Group_2_Value<-rep(NULL, 2*ItemNo)
      for (j in 1:ItemNo)
      {
        Value[2*j-1]<-etsclass_g[j]
        Value[2*j]<-etsclass_h[j]

        Group_1_Value[2*j-1]<-sum(1*(Data$gender[is.na(score_withna[,j])==F]=="F"))
        Group_1_Value[2*j]<-sum(1*(Data$heritage_t[is.na(score_withna[,j])==F]=="Y"))
        Group_2_Value[2*j-1]<-sum(1*(Data$gender[is.na(score_withna[,j])==F]=="M"))
        Group_2_Value[2*j]<-sum(1*(Data$heritage_t[is.na(score_withna[,j])==F]=="N"))
      }
      Group_1_Label<-rep(c("female","heritage"))
      Group_2_Label<-rep(c("male","non-heritage"))
      dif_csv<-cbind(item_id, collection_id, Type, Value, Group_1_Label, Group_1_Value, Group_2_Label, Group_2_Value)
      CsvPath<-paste0(MainPath,"/UploadFiles/",Skill,"/csv/")
      if (file.exists(CsvPath)==F)
      {
        dir.create(CsvPath)
      }
      write.csv(dif_csv, paste0(CsvPath, "diffs.csv"),row.names = F)

    }

  #### CTT analysis ####
  difficulty<-prop1<-colSums(score_matrix)/PerNo
  stdev<-stdev1<-apply(score_matrix,2,sd)
  # index for top 27% and bottom 27%
  n_27<-round(0.27*PerNo)
  top_27<-sort.int(sum_score,decreasing = T, index.return=T)$ix[1:n_27]
  bottom_27<-sort.int(sum_score,index.return=T)$ix[1:n_27]
  # compute discrimination with top and bottom 27, thus not equal to biserial correlation
  discrimination<-(colSums(score_matrix[top_27,])-colSums(score_matrix[bottom_27,]))/n_27
  # biserial correlation
  cor1<-cor2<-cor3<-cor4<-NULL
  for (j in 1:ItemNo)
  {
    if (sum(score_matrix[,j])==PerNo | sum(score_matrix[,j]==0)){
      cor1[j]<-0
    } else cor1[j]<-biserial(sum_score-score_matrix[,j],score_matrix[,j])
    if (sum(dist_1[,j])==PerNo | sum(dist_1[,j])==0) {
      cor2[j]<-0
    } else cor2[j]<-biserial(sum_score,dist_1[,j])
    if (sum(dist_2[,j])==PerNo | sum(dist_2[,j])==0) {
      cor3[j]<-0
    } else cor3[j]<-biserial(sum_score,dist_2[,j])
    if (sum(dist_3[,j])==PerNo | sum(dist_3[,j])==0) {
      cor4[j]<-0
    } else cor4[j]<-biserial(sum_score,dist_3[,j])
  }

  prop2<-colSums(dist_1)/PerNo
  stdev2<-apply(dist_1,2,sd)
  prop3<-colSums(dist_2)/PerNo
  stdev3<-apply(dist_2,2,sd)
  prop4<-colSums(dist_3)/PerNo
  stdev4<-apply(dist_3,2,sd)
  ctt_txt<-cbind(name,difficulty,stdev,discrimination,
                 prop1,stdev1,cor1,prop2,stdev2,cor2,
                 prop3,stdev3,cor3,prop4,stdev4,cor4)

  option_id<-rep(1:4,length=ItemNo*4)
  item_id<-rep(name,each=4)
  collection_id<-rep(TestName,length=ItemNo*4)
  Proportion<-c(rbind(prop1,prop2,prop3,prop4))
  Biserial_Corr<-c(rbind(cor1,cor2,cor3,cor4))
  Biserial_Corr[which(Biserial_Corr=="NaN")]<-0
  ctt_csv<-cbind(option_id,item_id,collection_id,Proportion,Biserial_Corr)
  write.csv(ctt_csv, paste0(CsvPath,"Options.csv"), row.names=F)

  #### IRT Analysis ####


  m<-RM(score_withna)
  p<-person.parameter(m)
  fit<-itemfit(p)
  bparam<-se<-wms<-stdwms<-ums<-stdums<-NULL
  for (j in 1:ItemNo)
  {
    if (prod(1*(na.omit(score_withna[,j])==1))==1)
    {
      bparam[j]<--4.20
      se[j]<-1.8345295
      wms[j]<-stdwms[j]<-ums[j]<-stdums[j]<-NA
    } else {
      if (prod(1*(na.omit(score_withna[,j])==0))==1)
      {
        bparam[j]<-4.20
        se[j]<-1.8345295
        wms[j]<-stdwms[j]<-ums[j]<-stdums[j]<-NA
      } else {
        ind_j<-which(colnames(m$X)==colnames(score_withna)[j])
        bparam[j]<-m$betapar[ind_j]*(-1)
        se[j]<-m$se.beta[ind_j]
        wms[j]<-fit$i.infitMSQ[ind_j]
        stdwms[j]<-fit$i.infitZ[ind_j]
        ums[j]<-fit$i.outfitMSQ[ind_j]
        stdums[j]<-fit$i.infitZ[ind_j]
      }
    }
  }

  model<-rep("L3",ItemNo)
  ncat<-rep(2,ItemNo)
  group<-name
  extreme<-rep("No",ItemNo)

  extreme[which(bparam==-4.20)]<-"Minimum"
  extreme[which(bparam==4.2)]<-"Maximum"

  # save to Reading or Listening file
  irt_full<-cbind(name,model,ncat,group,extreme,bparam,se,wms,stdwms,ums,stdums)
  IrtPath<-paste0(MainPath, "/", Skill, "/IRT/") # in Reading or Listening file
  if (file.exists(IrtPath)==F) 
    {
    dir.create(IrtPath)
    }
  write.csv(irt_full,paste0(IrtPath,Skill,"_IRT_Param.csv"),row.names = F)

  # save to Uploads/csv file
  Sample_size<-rep(PerNo,ItemNo)
  Response_Rate<-1-colSums(is.na(score_withna))/PerNo
  IRT_Infit<-wms
  IRT_Outfit<-ums
  IRT_Model<-rep("1PL",ItemNo)
  IRT_param<-bparam
  IRT_Standard_Error<-se
  irt_csv<-cbind(item_id,collection_id,Sample_size,Response_Rate,IRT_Infit,IRT_Outfit,IRT_Model,IRT_param,IRT_Standard_Error)
  write.csv(irt_csv,paste0(CsvPath,"items.csv"), row.names=F)

  #result[[which(c("Reading","Listening")==Skill)]]<-
  #  list(describe<-sum_describe, DIF_gender<-dif_full_g, DIF_heritage<-dif_full_h, CTT<-ctt_txt, IRT<-irt_full)
  }
  #result
}

