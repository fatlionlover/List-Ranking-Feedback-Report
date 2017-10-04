install.packages("openxlsx")
install.packages("plyr")
install.packages("pryr")
library("openxlsx")
library("plyr")
library(pryr)

mediumco2016=read.xlsx(file.choose(""))
mediumco=read.xlsx(file.choose(""))
medium2016winners=subset(mediumco2016,mediumco2016$Winners==1)
winners=subset(mediumco,mediumco$Winners==1)
##For All Plots
#2017

#clientname <- as.array(c("VTS","RevZilla.com","Intellinet","Inspirus, LLC","Frontline Performance Group","Amobee","DSLD Homes","Etsy","National Corporate Housing","Hagerty","Eventbrite","Wingstop Restaurants, Inc.","Katalyst Technologies","Elevate","VMS BioMarketing","Nextiva","Southern Pipe & Supply Company, Inc."))
#clientname <- as.array(c("4imprint, Inc.", "Granite Properties, Inc."))
clientname <- as.array(c("Greenleaf Trust","GoFundMe","TCG, Inc","Stellar Solutions, Inc.","Total Merchant Services","Yext","Child Trends","Bankers Healthcare Group","WillowTree, Inc.","Insomniac Games, Inc.","Credera","Shawmut Design and Construction","Integrated Project Management Company, Inc.","EKS&H LLLP","Health Catalyst","Edmunds","ENGEO Incorporated","Xactly Corporation","4imprint, Inc.", "Granite Properties, Inc."))
#i=1
for (i in 1:length(clientname)) {
  # cwd <- setwd("C://Users//nancy.cesena//Documents//SI&D//Credit Score//Plots//Medium")
  # newdir <- paste0("C://Users//nancy.cesena//Documents//SI&D//Credit Score//Plots//Medium//",clientname[1])
  # setwd(newdir)
  # 
mean4all=mean(mediumco$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', na.rm=TRUE); sd4all=sd(mediumco$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', na.rm=TRUE)
#lb=85.21; ub=95.42
x <- sort(mediumco$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', decreasing=TRUE)
hx <- dnorm(x,mean4all,sd4all)
#winners

winners=subset(mediumco,mediumco$Winners==1)
mean4allwinners=mean(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score',na.rm=TRUE); sd4allwinners=sd(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score',na.rm=TRUE)
xwinners <- sort(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', decreasing=TRUE)
hxwinners <- dnorm(xwinners,mean4allwinners,sd4allwinners)

# a %<a-% {
plot(xwinners, hxwinners, type="l", xlab="For All Scores", bty='n', ylab="Relative Frequency of Competitors",xlim=c(round(min(x),-1),mround(max(x),5)), yaxt='n',lwd=2,cex=2,
     main="", axes=TRUE, col="blue")
lines(x,hx, type="l", col="black")
abline(v=subset(mediumco,mediumco$Company_Name==clientname[i])$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', col="red")

# png(file=paste("C://Users//nancy.cesena//Documents//SI&D//Credit Score//Plots//Medium//",clientname[],"//forall.png"))
# a
# dev.off()

### Executive Team Effectiveness
#Size.Adjusted.Overall.Executive.Team.Effectiveness.Score
#2017
meanETE=mean(na.exclude(mediumco$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score')); sdETE=sd(na.exclude(mediumco$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'))
#lb=85.21; ub=95.42
ete <- sort(na.exclude(mediumco$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'), decreasing=TRUE)
hete <- dnorm(ete,meanETE,sdETE)

meanetewinners=mean(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score',na.rm=TRUE); sdetewinners=sd(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score',na.rm=TRUE)
etewinners <- sort(na.exclude(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'), decreasing=TRUE)
hetewinners <- dnorm(etewinners,meanetewinners,sdetewinners)


plot(etewinners, hetewinners, type="l", xlab="Executive Team Effectiveness Scores", bty='n',yaxt='n',xlim=c(round(min(ete),-1),mround(max(ete),5)),ylab="Relative Frequency of Competitors", lwd=2,cex=2,
     main="", axes=TRUE, col="blue")
lines(ete, hete, type="l", col="black")
abline(v=subset(mediumco,mediumco$Company_Name==clientname[i])$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', col="red")

# Innovation
#Size.Adjusted.Overall.Innovation.Score
meaninno=mean(na.exclude(mediumco$'Size.Adjusted.Overall.Innovation.Score')); sdinno=sd(na.exclude(mediumco$'Size.Adjusted.Overall.Innovation.Score'))
#lb=85.21; ub=95.42
inno <- sort(na.exclude(mediumco$'Size.Adjusted.Overall.Innovation.Score'), decreasing=TRUE)
hinno <- dnorm(inno,meaninno,sdinno)


meaninnowinners=mean(winners$'Size.Adjusted.Overall.Innovation.Score',na.rm=TRUE); sdinnowinners=sd(winners$'Size.Adjusted.Overall.Innovation.Score',na.rm=TRUE)
innowinners <- sort(na.exclude(winners$'Size.Adjusted.Overall.Innovation.Score'), decreasing=TRUE)
hinnowinners <- dnorm(innowinners,meaninnowinners,sdinnowinners)
#innovationplot
plot(innowinners, hinnowinners, type="l", xlab="Innovation Experience Scores", xlim=c(round(min(inno),-1),mround(max(inno),5)),bty='n',yaxt='n',ylab="Relative Frequency of Competitors", lwd=2,cex=2,
     main="", axes=TRUE, col="blue")
lines(inno, hinno, type="l", col="black")
abline(v=subset(mediumco,mediumco$Company_Name==clientname[i])$'Size.Adjusted.Overall.Innovation.Score', col="red")



}

# 
# 
# for (i in length(clientname)) {
  cwd <- setwd("C://Users//nancy.cesena//Documents//SI&D//Credit Score//Plots/Medium")
  newdir <- paste0(cwd, name1)
  setwd(newdir)
  moist<-raster(paste0("R://moist_tif/ind_moist",i,".tif"))
  writeRaster(moist,"moist.tif")
  setwd(cwd)
# }


