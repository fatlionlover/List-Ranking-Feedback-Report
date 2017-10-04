install.packages("openxlsx")
library("openxlsx")

mediumco2016=read.xlsx(file.choose(""))
mediumco=read.xlsx(file.choose(""))
medium2016winners=subset(mediumco2016,mediumco2016$Winners==1)
winners=subset(mediumco,mediumco$Winners==1)

#name1="PPR Talent Management Group"
mround <- function(x,base){ 
  base*round(x/base) 
} 

#Winners ForAll Plot
mean4allwinners=mean(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score') ; sd4allwinners=sd(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score')
#lb=85.21; ub=95.42
xwinners <- sort(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', decreasing=TRUE)
hxwinners <- dnorm(xwinners,mean4allwinners,sd4allwinners)
#winners 2016
mean4allwinners2016=mean(na.exclude(medium2016winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score')); sd4allwinners2016=sd(medium2016winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score',na.rm=TRUE)
#lb=85.21; ub=95.42
xwinners2016 <- sort(na.exclude(medium2016winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score'), decreasing=TRUE)
hxwinners2016 <- dnorm(xwinners2016,mean4allwinners2016,sd4allwinners2016)
#plot winners
plot(xwinners,hxwinners, xlab="For All Scores", type='l',xlim=c(round(min(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score',medium2016winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', na.rm=TRUE),1),
                                                                mround(max(winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score',medium2016winners$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', na.rm=TRUE),5)),
     bty="n", yaxt='n',ylab="Relative Frequency of Competitors", lwd=2,cex=2, main="", axes=TRUE)
lines(xwinners2016,hxwinners2016,type="l",col="blue")
# abline(v=subset(mediumco,mediumco$Company_Name==name1)$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', col="black", lty=2)
# abline(v=subset(mediumco2016,mediumco2016$Company_Name==name1)$'Size.Adjusted.For.All.(Experience.and.Diversity).Score', col="blue", lty=2)

#winners ete plots
meanetewinners=mean(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', na.rm=TRUE); sdetewinners=sd(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score',na.rm=TRUE)
#lb=85.21; ub=95.42
xetewinners <- sort(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', decreasing=TRUE)
hetewinners <- dnorm(xetewinners,meanetewinners,sdetewinners)
#winners 2016
meanetewinners2016=mean(medium2016winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', na.rm=TRUE); sdetewinners2016=sd(medium2016winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score',na.rm=TRUE)
#lb=85.21; ub=95.42
xetewinners2016 <- sort(na.exclude(medium2016winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'), decreasing=TRUE)
hetewinners2016 <- dnorm(xetewinners2016,meanetewinners2016,sdetewinners2016)
#plot winners
plot(xetewinners,hetewinners, xlab="Executive Team Effectiveness Scores", bty= 'n', type='l', xlim=c(round(min(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score',medium2016winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', na.rm=TRUE),-1),
                                                                                                         mround(max(winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score',medium2016winners$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', na.rm=TRUE),5)),
     yaxt='n',ylab="", lwd=2,cex=2, main="", axes=TRUE)
lines(xetewinners2016,hetewinners2016,type="l",col="blue")
# abline(v=subset(mediumco,mediumco$Company_Name==name1)$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', col="black", lty=2)
# abline(v=subset(mediumco2016,mediumco2016$Company_Name==name1)$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score', col="blue", lty=2)


#winners inno plots
meaninnowinners=mean(winners$'Size.Adjusted.Overall.Innovation.Score', na.rm=TRUE); sdinnowinners=sd(winners$'Size.Adjusted.Overall.Innovation.Score', na.rm=TRUE)
#lb=85.21; ub=95.42
xinnowinners <- sort(winners$'Size.Adjusted.Overall.Innovation.Score', decreasing=TRUE)
hinnowinners <- dnorm(xinnowinners,meaninnowinners,sdinnowinners)
#winners 2016
meaninnowinners2016=mean(na.exclude(medium2016winners$'Size.Adjusted.Overall.Innovation.Score')); sdinnowinners2016=sd(na.exclude(medium2016winners$'Size.Adjusted.Overall.Innovation.Score'))
#lb=85.21; ub=95.42
xinnowinners2016 <- sort(na.exclude(medium2016winners$'Size.Adjusted.Overall.Innovation.Score'), decreasing=TRUE)
hinnowinners2016 <- dnorm(xinnowinners2016,meaninnowinners2016,sdinnowinners2016)
#plot winners
plot(xinnowinners,hinnowinners, xlab="Innovation Experience Scores",bty='n', type='l',xlim=c(round(min(winners$'Size.Adjusted.Overall.Innovation.Score',medium2016winners$'Size.Adjusted.Overall.Innovation.Score', na.rm=TRUE),1),
                                                                                             mround(max(winners$'Size.Adjusted.Overall.Innovation.Score',medium2016winners$'Size.Adjusted.Overall.Innovation.Score', na.rm=TRUE),5)),
     yaxt='n',ylab="", lwd=2,cex=2, main="", axes=TRUE)
lines(xinnowinners2016,hinnowinners2016,type="l",col="blue")
# abline(v=subset(mediumco,mediumco$Company_Name==name1)$'Size.Adjusted.Overall.Innovation.Score', col="black", lty=2)
# abline(v=subset(mediumco2016,mediumco2016$Company_Name==name1)$'Size.Adjusted.Overall.Innovation.Score', col="blue",lty=2)




# 
# ##For All Plots
# mean4all=mean(na.exclude(mediumco$'Size.Adjusted.For.All.(Experience.and.Diversity).Score')); sd4all=sd(na.exclude(mediumco$'Size.Adjusted.For.All.(Experience.and.Diversity).Score'))
# lb=85.21; ub=95.42
# x <- sort(na.exclude(mediumco$'Size.Adjusted.For.All.(Experience.and.Diversity).Score'), decreasing=TRUE)
# hx <- dnorm(x,mean4all,sd4all)
# 
# #medium2016
# mean4all2016=mean(na.exclude(mediumco2016$'Size.Adjusted.For.All.(Experience.and.Diversity).Score')); sd4all2016=sd(na.exclude(mediumco2016$'Size.Adjusted.For.All.(Experience.and.Diversity).Score'))
# #lb=85.21; ub=95.42
# x2016 <- sort(na.exclude(mediumco2016$'Size.Adjusted.For.All.(Experience.and.Diversity).Score'), decreasing=TRUE)
# hx2016 <- dnorm(x2016,mean4all2016,sd4all2016)
# 
# plot(x, hx, type="l", xlab="For All Scores", ylab="Frequency", lwd=2,cex=2,
#      main="2016 vs 2017 For All Scores", axes=TRUE)
# #abline(v=c(85.21,95.42), h=c(.001,.001),col="black")
# 
# lines(x2016,hx2016,type="l",col="blue")
# 
# #2016 list
# cord.x <- c(80.498,seq(80.498,94.404,.02),94.404)
# cord.y <- c(0, dnorm(seq(80.498,94.404,.02)),0)
# 
# #abline(v=c(80.498,94.404), h=.01,col="blue")
# lines(cord.x,cord.y,col="blue")
# 
# ##2017 list
# cord.x <- c(85.21,seq(85.21,95.42,.01),95.42)
# cord.y <- c(0, dnorm(seq(85.21,95.42,.01)),0)
# 
# lines(cord.x,cord.y,col="black")
# 
# #lines(c(x[x==93.67 ],93.67), c(hx[x==93.67],0),col="black")
# 
# 
# ### Executive Team Effectiveness
# #Size.Adjusted.Overall.Executive.Team.Effectiveness.Score
# #2017
# meanETE=mean(na.exclude(mediumco$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score')); sdETE=sd(na.exclude(mediumco$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'))
# #lb=85.21; ub=95.42
# ete <- sort(na.exclude(mediumco$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'), decreasing=TRUE)
# hete <- dnorm(ete,meanETE,sdETE)
# #2016
# meanETE2016=mean(na.exclude(mediumco2016$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score')); sdETE2016=sd(na.exclude(mediumco2016$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'))
# #lb=85.21; ub=95.42
# ete2016 <- sort(na.exclude(mediumco2016$'Size.Adjusted.Overall.Executive.Team.Effectiveness.Score'), decreasing=TRUE)
# hete2016 <- dnorm(ete2016,meanETE2016,sdETE2016)
# 
# 
# plot(ete, hete, type="l", xlab="Executive Team Effectiveness Scores", ylab="Frequency", lwd=2,cex=2,
#      main=paste("2016 vs 2017 Executive Team Effectiveness Scores", "Amongst Competitors", sep="\n"), axes=TRUE)
# 
# lines(ete2016,hete2016, type="l", col="blue")
# 
# 
# #2016 range
# cord.x <- c(62.712,seq(62.712,91.33,.01),91.33)
# cord.y <- c(0, dnorm(seq(62.712,91.33,.01)),0)
# 
# lines(cord.x,cord.y,col="blue")
# 
# #2017 range
# cord.x <- c(69.52,seq(69.52,92.82,.01),92.82)
# cord.y <- c(0, dnorm(seq(69.52,92.82,.01)),0)
# 
# lines(cord.x,cord.y,col="black")
# 
# # Innovation
# #Size.Adjusted.Overall.Innovation.Score
# #2017
# meaninno=mean(na.exclude(mediumco$'Size.Adjusted.Overall.Innovation.Score')); sdinno=sd(na.exclude(mediumco$'Size.Adjusted.Overall.Innovation.Score'))
# #lb=85.21; ub=95.42
# inno <- sort(na.exclude(mediumco$'Size.Adjusted.Overall.Innovation.Score'), decreasing=TRUE)
# hinno <- dnorm(inno,meaninno,sdinno)
# #2016
# meaninno2016=mean(na.exclude(mediumco2016$'Size.Adjusted.Overall.Innovation.Score')); sdinno2016=sd(na.exclude(mediumco2016$'Size.Adjusted.Overall.Innovation.Score'))
# #lb=85.21; ub=95.42
# inno2016 <- sort(na.exclude(mediumco2016$'Size.Adjusted.Overall.Innovation.Score'), decreasing=TRUE)
# hinno2016 <- dnorm(inno2016,meaninno2016,sdinno2016)
# 
# plot(inno, hinno, type="l", xlab="Innovation Scores", ylab="Frequency", lwd=2,cex=2,
#      main=paste("2016 vs 2017 Innovation Scores", "Amongst Competitors",sep="\n"), axes=TRUE)
# 
# lines(inno2016,hinno2016, type="l", col="blue")
# 
# 
# #2016 range
# cord.x <- c(75.214,seq(75.214,92.918,.01),92.918)
# cord.y <- c(0, dnorm(seq(75.214,92.918,.01)),0)
# 
# lines(cord.x,cord.y,col="blue")
# 
# #2017 range
# cord.x <- c(80.48,seq(80.48,94.85,.01),94.85)
# cord.y <- c(0, dnorm(seq(80.48,94.85,.01)),0)
# 
# lines(cord.x,cord.y,col="black")
