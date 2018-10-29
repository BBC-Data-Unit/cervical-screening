#This script was used to grab the spreadsheets from https://www.gov.uk/government/publications/cervical-screening-coverage-and-data
#And then extract the data within those which was otherwise only accessible by using dropdown menus

install.packages("XLConnect")
install.packages("XLConnectJars")

#want to get locked data
require(XLConnect)

#import the workbook
wb = loadWorkbook ("Cervical_coverage_CCG_GPpracticeV2a_Dec2017.xlsx")

#get the titles of the sheets
getSheets(wb)

#create a dataframe just for the sheet called CCG, one for Dec2017 and one for Dec 2016
ccg = readWorksheet(wb, sheet = "CCG", header = TRUE)

Dec2017 = readWorksheet (wb, sheet = "Dec2017")

Dec2016 = readWorksheet (wb, sheet = "Dec16")


#write some CSVs to explore

write.csv(ccg, "ccgs.csv")

write.csv(Dec2017, "cervicaldec2017.csv")
write.csv(Dec2016, "cervicaldec2016.csv")

