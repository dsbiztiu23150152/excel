#演習課題
d <- read.csv('C:/Users/81801/Downloads/heights100.csv')
d

library(openxlsx)

wb <- buildWorkbook(d, 
                    asTable = TRUE, 
                    tableStyle = 'TableStyleLight8')

# 新規作成するExcelファイル（ワークブック）のファイルパス
FILE <- 'heights100.xlsx'

saveWorkbook(wb, file = FILE, overwrite = TRUE)

addWorksheet(wb, sheetName = 'Sheet 2', gridLines = FALSE)

#p1 = hist(d$male)
#p2 = hist(d$female)

#insertPlot(wb, sheet = 'Sheet 2',width = 10, height = 10)

hist(d$male, main = "Male Height", xlab = "Height")
insertPlot(wb, sheet = "Sheet 2", width = 6, height = 4, startCol = 1, startRow = 1)

hist(d$female, main = "Female Height", xlab = "Height")
insertPlot(wb, sheet = "Sheet 2", width = 6, height = 4, startCol = 10, startRow = 1)

writeDataTable(wb, sheet = 'Sheet 2', x = d, 
               tableStyle = 'TableStyleLight8', startCol = 1, startRow = 22)

saveWorkbook(wb, file = FILE, overwrite = TRUE)

openXL('heights100.xlsx')