

install.packages("tabulizer")
install.packages("dplyr")
install.packages("rJava")
install.packages("devtools")
install.packages("pdftools")
install.packages("tidyverse")
install.packages("glue")
install.packages("tabularizejars")


if (!require('knitr')) {
  install.packages("knitr")
}
if (!require('devtools')) {
  install.packages("devtools")
}
if (!require('RWordPress')) {
  devtools::install_github(c("duncantl/XMLRPC", "duncantl/RWordPress"))
}


library(RWordPress)
library(knitr)

# Tell RWordPress how to set the user name, password, and URL for your WordPress site.
options(WordpressLogin = c(jwildish = "Hazel_1912"),
        WordpressURL = "https://yot.hnq.mybluehost.me/xmlrpc.php")

?knit2wp

# If need be, set your working directory to the location where you stored the Rmd file. 
setwd("C:/Jordan/Desktop/RWD")
getwd()
# Tell knitr to create the html code and upload it to your WordPress site
knit2wp('RmdTest.rmd', title = 'Blog Posting from R Markdown to WordPress',publish = FALSE)


#Step 1 Adobe PDF to Excel Conversion

#Step 2 Import 
library(readxl)
arb_offset_credit_issuance_table <- read_excel("C:/Users/Jordan/Desktop/E2/arb_offset_credit_issuance_table.xlsx", 
                                               skip = 4)
View(arb_offset_credit_issuance_table)

names(arb_offset_credit_issuance_table)

arb_offset_credit_issuance_table$X__1 <- NULL
arb_offset_credit_issuance_table$X__2 <- NULL
arb_offset_credit_issuance_table$X__3 <- NULL
arb_offset_credit_issuance_table$X__4 <- NULL
arb_offset_credit_issuance_table$X__5 <- NULL
arb_offset_credit_issuance_table$X__6 <- NULL

arb_offset_credit_issuance_table <- subset(arb_offset_credit_issuance_table, 'Project Name' != `Project Name`)



# download CSV from https://thereserve2.apx.com/myModule/rpt/myrpt.asp?r=111
library(dplyr)

CARMerge <- read.csv("C:/Users/Jordan/Desktop/E2/temp.csv")

CARMerge$Project.Registered.Date <- as.Date(CARMerge$Project.Registered.Date,"%m/%d/%Y")

CARMerge2 <- CARMerge %>% filter(Project.Registered.Date > 2017-01-01)


#additional information available at https://thereserve2.apx.com/myModule/rpt/myrpt.asp

arbcrossreference <- read.csv("C:/Users/Jordan/Desktop/E2/Offsets.csv")





#ARB table creator (from table made in acrobat)

arb_offset_credit_issuance_table <- read_excel("C:/Users/Jordan/Desktop/arb_offset_credit_issuance_table.xlsx", 
                                               skip = 4)
install.packages("kableExtra")
library(kableExtra)
library(dplyr)
install.packages("devtools")
library(devtools)
devtools::install_github("haozhu233/kableExtra")
names(arb_offset_credit_issuance_table)

arb_offset_credit_issuance_table <- arb_offset_credit_issuance_table %>% select(`ARB Project ID #`, `Project Name`, `Type of Protocol`, `Offset Project Operator`, `Project Type`, `Project Documentation` )

arb_offset_credit_issuance_table <- subset(arb_offset_credit_issuance_table, `Project Name` != "Project Name")

table(arb_offset_credit_issuance_table$`Offset Project Operator`)

kable(arb_offset_credit_issuance_table) %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))

kable(arb_offset_credit_issuance_table, align = "c") %>% 
  kable_styling(full_width =  F) %>% 
  column_spec(1, bold = T) %>%
  collapse_rows(columns= 3:4, valign = "top")







collapse_rows_dt <- data.frame(C1 = c(rep("a", 10), rep("b", 5)),
                               C2 = c(rep("c", 7), rep("d", 3), rep("c", 2), rep("d", 3)),
                               C3 = 1:15,
                               C4 = sample(c(0,1), 15, replace = TRUE))

kable(collapse_rows_dt, "latex", booktabs = T, align = "c") %>%
  column_spec(1, bold=T) %>%
  collapse_rows(columns = 1:2, latex_hline = "major", valign = "middle") %>% kable_styling()
