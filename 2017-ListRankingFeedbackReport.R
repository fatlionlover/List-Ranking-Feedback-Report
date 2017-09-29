#List Ranking Feedback Report

library(svGUI); library(svDialogs); library(openxlsx); library(plyr)
library(RDCOMClient); library(ReporteRs); library(officer); library(grid)
library(XML);library(grImport);library(magrittr)
library(rJava);library(ggplot2);library(reshape2);
library(scales);library(dplyr); library(doBy); library(forcats)
library(personograph)

lrfrtemp <- "C:/Users/chandni.kazi/Desktop/List Ranking Feedback Report/Best Medium Workplaces Template.docx"
lrfrfinal <- "C:/Users/chandni.kazi/Desktop/List Ranking Feedback Report/Best Medium Workplaces FINAL.docx"
mydoc <- docx(template =lrfrtemp)
styles(mydoc)
list_bookmarks(mydoc)

clientname <- "Wazzup Company"
pg2_table_2017rank <- 10
if (pg2_table_2017rank>100){
  estimated2017<-"Estimated"
} else {
  estimated2017 <- ""
}
pg2_table_2016rank <- 99
if (pg2_table_2016rank>100){
  estimated2016<-"Estimated"
} else {
  estimated2016 <- ""
}
pg2_table_rankchange <- pg2_table_2016rank - pg2_table_2017rank
if (estimated2017=="Estimated" | estimated2016=="Estimated"){
  estimated2017<-"Estimated"
} else {
  estimated2017 <- ""
}


mydoc %>%
  addParagraph( value = clientname, stylename = "titleclientname", bookmark = "ClientName" ) %>%
  addParagraph( value = clientname, stylename = "clientnamestyle", bookmark = "ClientName" ) %>%
  addParagraph( value = as.character(pg2_table_2017rank), stylename = "summarytable", bookmark = "pg2_table_2017rank" ) %>%
  addParagraph( value = as.character(pg2_table_2016rank), stylename = "summarytable", bookmark = "pg2_table_2016rank" ) %>%
  addParagraph( value = as.character(pg2_table_rankchange), stylename = "summarytable", bookmark = "pg2_table_rankchange" ) %>%
  addParagraph( value = "John Doe", stylename = "small", bookmark = "pg3_amongbest" ) %>%
  addParagraph( value = "John Doe", stylename = "small", bookmark = "pg3_amongcertified" ) %>%
  addPlot(mydoc,fun=function() hist(testdata)) %>%
  addParagraph( value = "John Doe", stylename = "small", bookmark = "pg4_amongbest" ) %>%
  addParagraph( value = "John Doe", stylename = "small", bookmark = "pg4_amongcertified" ) %>%
  addParagraph( value = "John Doe", stylename = "small", bookmark = "pg4_ETE_histogram" ) %>%
  addFlexTable( flextable = ft, bookmark = "DATA" ) %>%
  addPlot( fun = print, x = myplot1, bookmark = "PLOT" ) %>%
  writeDoc( file = lrfrfinal)



testing <- "this is a test"
testnum <- 23
testdata <- list(first=0.9, second=0.1)

mydoc <- addPageBreak(mydoc, )
mydoc <- addTitle(mydoc,"Testing: Your List Ranking Component Details",level = 1)
mydoc <- addPlot(mydoc,fun=function() personograph(testdata))

writeDoc(mydoc,lrfrfinal)
