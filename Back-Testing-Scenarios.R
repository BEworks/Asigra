rm(list=ls(all=TRUE))
for(k in 2013:2015){
  rm(list=ls(all=TRUE)[!ls(all=TRUE)%in%c("ROMIstats13","ROMIstats14","k")])
  # Back Testing Routine
  
  # packages
  library(xlsx)
  library(plyr)
  
  # data imports ----
  directory<-"C://Users//BE//Dropbox (BEworks)//Our Client work//Asigra//Compensation Plan//Mandate 4 - Marketing Team//Marketing Metrics//Modelling//ROMI Proper"
  setwd(directory)
  
  # import check for NA functions
  
  ## Financial Data
  setwd(paste0(directory,"//FinancialData"))
  
  finDat10m<-read.xlsx("FinancialData.xlsx",sheetName="<10m",colClasses=rep("numeric",4))
  finDat1020m<-read.xlsx("FinancialData.xlsx",sheetName="10-20m",colClasses=rep("numeric",4))
  finDat20m<-read.xlsx("FinancialData.xlsx",sheetName=">20m",colClasses=rep("numeric",4))
  markSpend<-read.xlsx("FinancialData.xlsx",sheetName="Spending")
  betas<-read.xlsx("FinancialData.xlsx",sheetName="Beta")
  discount<-read.xlsx("FinancialData.xlsx",sheetName="Discount")
  
  Rev2015<-read.xlsx("RevenueData.xlsx",sheetName="Revenue 2015")
  Rev2014<-read.xlsx("RevenueData.xlsx",sheetName="Revenue 2014")
  Rev2013<-read.xlsx("RevenueData.xlsx",sheetName="Revenue 2013")
  samples<-read.xlsx("RevenueData.xlsx",sheetName="Sample Size Data")
  
  
  ## Opportunities Data
  setwd(paste0(directory,"//OpportunitiesData"))
  
  opports<-read.xlsx("AllOpportunities.xlsx",sheetName="AllOpportunities",stringsAsFactors=F)
  opports<-within(opports,
                  {Created.Date<-as.Date(as.numeric(Created.Date)-2,origin="1900-01-01")
                  Close.Date<-as.Date(as.numeric(Close.Date)-2,origin="1900-01-01")})
  
  stages<-read.xlsx("AllOpportunities.xlsx",sheetName="Stages",stringsAsFactors=F,header=F)
  
  endOf14<-read.xlsx("BacktestData.xlsx",sheetName="2014 Leads",stringsAsFactors=F)
  colnames(endOf14)[2]<-"EndYear"
  endOf13<-read.xlsx("BacktestData.xlsx",sheetName="2013 Leads",stringsAsFactors=F)
  colnames(endOf13)[2]<-"EndYear"
  
  setwd(directory)
  
  # input values for backtest or scenarios ----
  
  # other
  
  scenarios<-T
  
  if(scenarios){
    k<-2016;
    numSmall<-28; numMed<-10; numLarge<-5
    revSmall<-5000; revMed<-15000; revLarge<-49000
    groSmall<-0.55; groMed<-0.3; groLarge<-0.5
    
    finDat10m[nrow(finDat10m)+1,]<-list(k,groSmall,revSmall,finDat10m[nrow(finDat10m),ncol(finDat10m)])
    finDat1020m[nrow(finDat1020m)+1,]<-list(k,groMed,revMed,finDat1020m[nrow(finDat1020m),ncol(finDat1020m)])
    finDat20m[nrow(finDat20m)+1,]<-list(k,groLarge,revLarge,finDat20m[nrow(finDat20m),ncol(finDat20m)])
    k<-k-1
    finDat10m[,1]<-finDat1020m[,1]<-finDat20m[,1]<-c(2006:2015)
  }
  BT<-T
  year<-k
  startDate<-paste0(year,"-01-01")
  endDate<-paste0((year+1),"-01-01")
  distYears<-10
  
  # Calculate acquisition process advancement probabilities ----
  
  ## generate table of stages
  genStages<-stages[1:6,1]
  stages<-stages[1:6,-1]
  stageList<-setNames(split(stages,seq(nrow(stages))),genStages)
  stageList<-lapply(stageList,function(x) x[!is.na(x)])
  
  relevantStages<-unlist(stageList)[unlist(stageList)%in%opports$Last.Stage]
  numStages<-length(relevantStages)
  
  ## build into table
  continueProb<-data.frame(stage=relevantStages,prob=numeric(length=numStages),cumProb=numeric(length=numStages),row.names=NULL)
  
  ### prob of advancing to the next stage
  for (i in 1:numStages){
    stage<-relevantStages[i]
    abandonCount<-sum(c(opports$Last.Stage==stage),na.rm=T)
    futureStages<-relevantStages[(i+1):numStages]
    continueCount<-sum((opports$Stage=="Closed Won"),na.rm=T)+sum((opports$Last.Stage%in%futureStages),na.rm=T)
    total<-abandonCount+continueCount
    continueProb$prob[i]<-continueCount/total
  }
  
  ### prob of becoming a customer
  for (i in 1:numStages){
    continueProb$cumProb[i]<-prod(continueProb$prob[i:numStages]) 
  }
  
  # calculate average time to revenue ----
  
  
  opports$Age<-opports$Close.Date-opports$Created.Date
  opports$Age[opports$Age<0]<-NA
  
  revDelay<-tapply(opports$Age,opports$Revenue.Size,mean,na.rm=T)
  revDelay<-revDelay[names(revDelay)!="NA"]
  
  # calc CLV -----
  
  listFinDat<-list("10m"=finDat10m,"1020m"=finDat1020m,"20m"=finDat20m)
  
  ## calc rolling rates
  calcRolling<-function(listFinDat,betas){
    
    for(i in 1:length(listFinDat)){
      # through profiles
      data<-listFinDat[[i]]
      aData<-data
      
      for(j in 2:ncol(data)){
        # through columns
        name<-colnames(data)[j]
        relBeta<-betas$Beta[betas[,1]==gsub("[.]"," ",name)]
        newName<-paste0("rolling.",name)
        aData[,newName]<-numeric(length=nrow(data))
        
        for(k in 1:nrow(data)){
          # through observations
          
          currK<-aData[k,name]
          if(!is.na(currK)){
            if(k==1){
              prevRollK<-currK
            }else{
              prevRollK<-rollK
            }
            
            rollK<-prevRollK+relBeta*(currK-prevRollK) 
          }else{
            rollK<-prevRollK
          }
          
          aData[k,newName]<-rollK
          
        }
      }
      listFinDat[[i]]<-aData
    }  
    return(listFinDat)
  }
  listFinDat<-calcRolling(listFinDat,betas)
  
  ## calc CLV for 3 buckets
  calcCLV<-function(listFinDat,discount,year){
    
    n=c(0,0,0)
    yn<-data.frame(y0=n,y1=n,y2=n,y3=n,y4=n,y5=n,y6=n,y7=n,y8=n,y9=n)
    CLV<-data.frame(profile=c("<10m","10-20m",">20m"),CLV=numeric(length=3),yn,stringsAsFactors=F)
    
    for(i in 1:length(listFinDat)){
      data<-listFinDat[[i]]
      data<-data[data[,1]==year,]
      
      growth<-data[,"rolling.Growth.Rate"]
      avgRev<-data[,"rolling.Average.Revenue.of.New"]
      aban<-data[,"rolling.Abandonment.Rate"]
      expRev=0
      for(j in 0:(distYears-1)){
        CLV[i,paste0("y",j)]<-(avgRev*(1+growth)^j*(1-aban)^j)/(1+discount)^j
      }
      CLV[i,"CLV"]<-rowSums(CLV[i,c(3:ncol(CLV))])
    }
    return(CLV)
  }
  CLV<-calcCLV(listFinDat,discount,year)
  
  ## functional ratio for calculating <20m bucket CLV
  smallshare<-samples[samples[,1]==year,]
  smallshare<-smallshare[,2]/(smallshare[,2]+smallshare[,3])
  ### add bucket clv
  less20<-list("<20m")
  for(i in 2:ncol(CLV)){
    less20[i]<-smallshare*CLV[1,i]+(1-smallshare)*CLV[2,i]
  }
  CLV[4,]<-less20
  
  # back test ----
  ## get appropriate backtest data set
  if(BT==T){
    
    
    BTopps<-opports[opports$Created.Date>startDate & opports$Created.Date<endDate & !is.na(opports$Created.Date),]
    BTopps<-merge(BTopps,endOf14,by="Account.Name",all.x=T)
    BTopps$EndYear[BTopps$Close.Date<endDate]<-NA
    BTopps$Stage[!is.na(BTopps$EndYear)]<-BTopps$EndYear[!is.na(BTopps$EndYear)]
    
  }
  
  ## asign probabilities to leads
  
  BTopps<-merge(BTopps,continueProb[,c(1,3)],by.x="Stage",by.y="stage",all.x=T)
  BTopps<-within(BTopps,{
    cumProb[Stage=="Closed Won"]<-1
    cumProb[Stage%in%c("Closed Lost","Closed No Decision")]<-0
  })
  
  ## assign CLV to leads
  
  estimateDate<-function(revDelay,BTopps,endDate){
    BTopps$Exp.Close<-BTopps$Close.Date
    for(i in 1:length(revDelay)){
      BTopps$Exp.Close[BTopps$Revenue.Size==names(revDelay)[i]]<-BTopps$Created.Date[BTopps$Revenue.Size==names(revDelay)[i]]+revDelay[i]
    }
    BTopps<-within(BTopps,{
      Exp.Close[Exp.Close<endDate]<-endDate
      Exp.Close[Close.Date<endDate]<-Close.Date[Close.Date<endDate]
    })
    BTopps$time2Rev<-BTopps$Exp.Close-as.Date(endDate)
    BTopps$time2Rev[BTopps$time2Rev<=0]<-0
    
    return(BTopps) 
  }
  BTopps<-estimateDate(revDelay,BTopps,endDate)
  
  if(scenarios){
    meanTime<-as.numeric(mean(BTopps$Close.Date,na.rm=T)-as.Date(endDate))
    BTopps<-BTopps[BTopps$cumProb>0 & BTopps$cumProb<1 & !is.na(BTopps$cumProb),]
    revSize=c(rep("<10m",numSmall),rep("10-20m",numMed),rep(">20m",numLarge)) 
    closed<-data.frame(Revenue.Size=revSize,time2Rev=rep(meanTime,length(revSize)),cumProb=rep(1,length(revSize))) #time2Rev is numerical, i have date
    BTopps<-rbind.fill(BTopps,closed)
  }
  
  assignCLV<-function(BTopps,CLV){
    delaycoeff<-(1-as.numeric(BTopps$time2Rev)/3650)
    BTopps<-merge(BTopps,CLV,by.x="Revenue.Size",by.y="profile",all.x=T)
    BTopps$expRev<-BTopps$CLV*BTopps$cumProb*delaycoeff
    return(BTopps)
  }
  
  
  
  
  BTopps<-assignCLV(BTopps,CLV)
  
  # calc final ROMI ------------
  
  ## calc expectation adjustment coefficient
  returns<-sum(BTopps$expRev,na.rm=T)
  revExp<-sum(BTopps$y0[BTopps$cumProb==1],na.rm=T)
  
  ### actual revenue
  if(year==2013){
    realRevs<-rowSums(Rev2013[1,2:4])
  }else if(year==2014){
    realRevs<-rowSums(Rev2014[1,2:4])
  }else if(year==2015){
    realRevs<-rowSums(Rev2015[1,2:4])
    theoRevs<-sum(samples[samples$NA.==year,2:4]*c(finDat10m[finDat10m$NA.==(year-1),3],finDat1020m[finDat1020m$NA.==(year-1),3],finDat20m[finDat20m$NA.==(year-1),3]))
    if(scenarios){
      realRevCoeff<-realRevs/theoRevs
      theoRevs2<-sum(numSmall*revSmall,numMed*revMed,numLarge*revLarge)
      realRevs<-theoRevs2*realRevCoeff
    }
  }
  
  
  ### coef
  adjCoef<-realRevs/revExp
  ### adjusted returns
  adjReturns<-returns*adjCoef
  
  # calc "actual ROMI" ------------
  
  calcReturn<-function(listFinDat,discount,year,Rev20){
    
    knownCount<-nrow(Rev20)
    n=c(0,0,0)
    yn<-data.frame(y0=n,y1=n,y2=n,y3=n,y4=n,y5=n,y6=n,y7=n,y8=n,y9=n)
    CLV<-data.frame(profile=c("<10m","10-20m",">20m"),CLV=numeric(length=3),yn,stringsAsFactors=F)
    
    for(i in 1:length(listFinDat)){
      data<-listFinDat[[i]]
      data<-data[data[,1]==year,]
      
      growth<-data[,"rolling.Growth.Rate"]
      Rev<-Rev20[knownCount,(i+1)]
      aban<-data[,"rolling.Abandonment.Rate"]
      for(j in 0:(distYears-1)){
        if(j<(knownCount)){
          revValue<-Rev20[(j+1),(i+1)]
        }else{
          revValue<-(Rev*(1+growth)^(j+1-knownCount)*(1-aban)^(j+1-knownCount))
        }
        CLV[i,paste0("y",j)]<-revValue/(1+discount)^j
      }
      CLV[i,"CLV"]<-rowSums(CLV[i,c(3:ncol(CLV))])
    }
    return(CLV)
  }
  if(year==2014){
    CLVreturns<-calcReturn(listFinDat,discount,year,Rev2014)
  }else if(year==2013){
    CLVreturns<-calcReturn(listFinDat,discount,year,Rev2013)
  }else if(year==2015){
    CLVreturns<-calcReturn(listFinDat,discount,year,Rev2015)
  }
  returns2<-sum(CLVreturns[,2])
  
  #output ------------
  investment<-rowSums(markSpend[markSpend[,1]==year,2:3])
  
  if(scenarios){
    investment<-3250000
  }
  
  ROMIstats<-data.frame(type=c("base","adjusted","actual"),Revenue=c(returns,adjReturns,returns2),year=year,investment=investment,coef=adjCoef)
  ROMIstats$ROMI<-ROMIstats[,2]/ROMIstats[,4]
  
  if(year=="2013"){
    ROMIstats13<-ROMIstats
  }else if(year=="2014"){
    ROMIstats14<-ROMIstats
  }else if(year=="2015"){
    ROMIstats15<-ROMIstats
  }
  
  #after -------
}

ROMIstats13
ROMIstats14
ROMIstats15


write.table(ROMIstats13, "clipboard", sep="\t", row.names=FALSE)
write.table(ROMIstats14, "clipboard", sep="\t", row.names=FALSE)

parts2015<-opports[opports$Created.Date>="2015-01-01" & opports$Close.Date<"2016-01-01" & opports$Stage=="Closed Won",c("Account.Name","Close.Date")]
write.table(parts2015, "clipboard", sep="\t", row.names=FALSE)




test<-data.frame(red=c(1,2),blue=c(3,4),green=c(5,6))
test2<-data.frame(red=c(7,8),green=c(9,10))
rbind.fill(test,test2)

