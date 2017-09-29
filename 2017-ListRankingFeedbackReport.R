#List Ranking Feedback Report
library(svGUI); library(svDialogs); library(openxlsx); library(plyr)
library(RDCOMClient); library(ReporteRs); library(officer); library(grid)
library(XML);library(grImport);library(magrittr)
library(rJava);library(ggplot2);library(reshape2);
library(scales);library(dplyr); library(doBy); library(forcats)
library(personograph)
#####################################################
#                  LRFR Template                    #
#####################################################

lrfrtemp <- "C:/Users/chandni.kazi/Desktop/List Ranking Feedback Report/Best Medium Workplaces Template.docx"

#####################################################
#             Name of LRFR for Client               #
#####################################################

lrfrfinal <- "C:/Users/chandni.kazi/Desktop/List Ranking Feedback Report/Best Medium Workplaces FINAL.docx"

#####################################################
#              Initialzing template to              #
#        replace bookmarks with client's data       #
#####################################################

mydoc <- docx(template =lrfrtemp)

#####################################################
#    Styles and bookmarks available in template     #
#####################################################

#styles(mydoc)
#list_bookmarks(mydoc)


#####################################################
#       DELETE THIS SECTION WHEN WE DON'T NEED      #
#                 DUMMY DATA ANYMORE                #
#####################################################

c <- 1
d <- 2
a <- -2
b <- 3.5
ll <- pnorm(a, c, d)
ul <- pnorm(b, c, d)
x <- qnorm( runif(3000, ll, ul), c, d )

#####################################################
#         Connecting Client's data to the           #
#        bookmarks in the MS word template          #
#####################################################

#Source([NANCY's R SCRIPT PATH HERE])

clientname <- "Wazzup Company"       #NANCY VARIABLE REPLACEMENT FROM HER SCRIPT  

#Page 2 
pg2_table_2017rank <- 10     #NANCY VARIABLE REPLACEMENT 
pg2_table_2016rank <- 99     #NANCY VARIABLE REPLACEMENT

if (pg2_table_2017rank > 100){
  estimated2017 <- "Estimated"
  estimatedchange <- "Estimated"
} else {
  estimated2017 <- " "
  estimatedchange <- " "
}

if(is.numeric(pg2_table_2016rank - pg2_table_2017rank)){
  pg2_table_rankchange <- pg2_table_2016rank - pg2_table_2017rank
} else {
  pg2_table_rankchange <- "N/A"
  estimatedchange <- " "
}


#Page 3 
pg3_amongbest <- paste0("better than ",10,"%")         #NANCY VARIABLE REPLACEMENT 
pg3_amongcertified <- paste0("better than ",20,"%")    #NANCY VARIABLE REPLACEMENT 
testplot1 <- plot(density(x))                           #NANCY PLOT REPLACEMENT 

#Page 4
pg4_amongbest <- paste0("better than ",20,"%")         #NANCY VARIABLE REPLACEMENT 
pg4_amongcertified <- paste0("better than ",30,"%")    #NANCY VARIABLE REPLACEMENT 
testplot2 <- plot(density(x))                           #NANCY PLOT REPLACEMENT 



#Page 5 
pg5_amongbest <- paste0("better than ",30,"%")         #NANCY VARIABLE REPLACEMENT 
pg5_amongcertified <- paste0("better than ",40,"%")    #NANCY VARIABLE REPLACEMENT 
testplot3 <- print(plot(density(x)))                           #NANCY PLOT REPLACEMENT 


#####################################################
#      Updating template to be client specific      #
#####################################################

mydoc %>%
  addParagraph( value = clientname, stylename = "titleclientname", bookmark = "ClientName" ) %>%
  addParagraph( value = clientname, stylename = "clientnamestyle", bookmark = "ClientName" ) %>%
  
  addParagraph( value = as.character(pg2_table_2017rank), stylename = "summarytable", bookmark = "pg2_table_2017rank" ) %>%
  addParagraph( value = as.character(pg2_table_2016rank), stylename = "summarytable", bookmark = "pg2_table_2016rank" ) %>%
  addParagraph( value = as.character(pg2_table_rankchange), stylename = "summarytable", bookmark = "pg2_table_rankchange" ) %>%
  addParagraph( value = as.character(estimated2017), stylename = "estimated", bookmark = "estimated2017" ) %>%
  addParagraph( value = as.character(estimatedchange), stylename = "estimated", bookmark = "estimatedchange" ) %>%
  
  addParagraph( value = clientname, stylename = "betterthantext", bookmark = "pg3_forall_histogram_clientname" ) %>%
  addParagraph( value = pg3_amongbest, stylename = "betterthantext", bookmark = "pg3_amongbest" ) %>%
  addParagraph( value = pg3_amongcertified, stylename = "betterthantext", bookmark = "pg3_amongcertified" ) %>%
  addPlot(mydoc,fun=function() plot(density(x),bookmark="pg3_forall_histogram" )) %>%
  
  addParagraph( value = clientname, stylename = "betterthantext", bookmark = "pg4_ETE_histogram_clientname" ) %>%
  addParagraph( value = pg4_amongbest, stylename = "betterthantext", bookmark = "pg4_amongbest" ) %>%
  addParagraph( value = pg4_amongcertified, stylename = "betterthantext", bookmark = "pg4_amongcertified" ) %>%
  addPlot(mydoc,fun=function() plot(density(x),bookmark="pg4_ETE_histogram" )) %>%
  
  addParagraph( value = clientname, stylename = "betterthantext", bookmark = "pg5_IE_histogram_clientname" ) %>%
  addParagraph( value = pg5_amongbest, stylename = "betterthantext", bookmark = "pg5_amongbest" ) %>%
  addParagraph( value = pg5_amongcertified, stylename = "betterthantext", bookmark = "pg5_amongcertified" ) %>%
  addPlot(mydoc,fun=function() plot(density(x),bookmark="pg5_IE_histogram" )) %>%
  writeDoc( file = lrfrfinal)

#####################################################
#        Convert Final Word Document to PDF         #
#####################################################






#####################################################
#              DELTE CODE BELOW AT END              #
#####################################################

# addParagraph( value = "John Doe", stylename = "small", bookmark = "pg4_amongbest" ) %>%
#   addParagraph( value = "John Doe", stylename = "small", bookmark = "pg4_amongcertified" ) %>%
#   addParagraph( value = "John Doe", stylename = "small", bookmark = "pg4_ETE_histogram" ) %>%
#   addFlexTable( flextable = ft, bookmark = "DATA" ) %>%

# addPlot(mydoc,fun=function() plot(density(x))) %>%
#   addPlot( fun = print, x = myplot1, bookmark = "PLOT" ) %>% 






