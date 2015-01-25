

require("openxlsx")
require('testthat')

# 
# 
# context("Re-ordering worksheets.")
# 
# 
# 
# test_that("Worksheet ordering from new Workbook", {
#   
  options('stringsAsFactors' = FALSE)
  tempFile <- file.path(tempdir(), "break.xlsx")
  
  df1 <- iris[1:5, 1:4]
  df2 <- mtcars
  
  
  df3 <- data.frame("Date" = Sys.Date()-0:10,
                   "Logical" = sample(c(TRUE, FALSE), 1, replace = TRUE),
                   "Currency" = as.numeric(-5:5)*100,
                   "Accounting" = as.numeric(-5:5),
                   "hLink" = "http://cran.r-project.org/", 
                   "Percentage" = seq(-5, 5, length.out=11),
                   "TinyNumber" = runif(11) / 1E9, stringsAsFactors = FALSE)
    
  df3U <- df3
    
  class(df3$Currency) <- "currency"
  class(df3$Accounting) <- "accounting"
  class(df3$hLink) <- "hyperlink"
  class(df3$Percentage) <- "percentage"
  class(df3$TinyNumber) <- "scientific"
  

  df4 <- data.frame("X" = 1:10000, "Y" = sample(LETTERS, size = 10000, replace = TRUE))
  df5 <- USJudgeRatings
    
  hs <- createStyle(fontColour = "blue", textRotation = 45)
  
  
  wb <- createWorkbook()
  expect_equal(names(wb), character(0))
  
  addWorksheet(wb = wb, sheetName = "Sheet 1", gridLines = FALSE, tabColour = "red", zoom = 75)
  writeDataTable(wb, sheet = 1, x = df1, startCol = 7, startRow = 10, tableName = "Sheet1Table1")
  expect_equal(names(wb),  "Sheet 1")
  
  
  addWorksheet(wb, sheetName = "Sheet 2", tabColour = "purple")
  writeDataTable(wb, sheet = "Sheet 2", x = df2, startCol = 2, startRow = 2, rowNames = TRUE)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2"))
  
  

  addWorksheet(wb, sheetName = "Sheet 3", tabColour = "green")
  writeDataTable(wb, sheet = 3, x = df3, startCol = 1, startRow = 1)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3"))
  
  addWorksheet(wb, sheetName = "Sheet 4", tabColour = "orange")
  writeDataTable(wb, sheet = 4, x = df4)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
  
  addWorksheet(wb, sheetName = "Sheet 5", tabColour = "yellow")
  writeData(wb, sheet = "Sheet 5", x = df5, rowNames = TRUE)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5"))
  

  
  worksheetOrder(wb) <- c(1, 3, 5, 4, 2)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5"))
  
  ## save and load 1
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)
  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 3", "Sheet 5", "Sheet 4", "Sheet 2"))

  expect_equal(df1,  read.xlsx(wb, sheet = 1))
  expect_equal(df1,  read.xlsx(wb, sheet = "Sheet 1"))
  expect_equal(df1,  read.xlsx(tempFile, sheet = 1))
  expect_equal(df1,  read.xlsx(tempFile, sheet = "Sheet 1"))
  
  
  expect_equal(df3U,  read.xlsx(wb, sheet = 2, detectDates = TRUE))
  expect_equal(df3U,  read.xlsx(wb, sheet = "Sheet 3", detectDates = TRUE))
  expect_equal(df3U,  read.xlsx(tempFile, sheet = 2, detectDates = TRUE))
  expect_equal(df3U,  read.xlsx(tempFile, sheet = "Sheet 3", detectDates = TRUE))
    
  
  expect_equal(df5,  read.xlsx(wb, sheet = 3, rowNames = TRUE))
  expect_equal(df5,  read.xlsx(wb, sheet = "Sheet 5", rowNames = TRUE))
  expect_equal(df5,  read.xlsx(tempFile, sheet = 3, rowNames = TRUE))
  expect_equal(df5,  read.xlsx(tempFile, sheet = "Sheet 5", rowNames = TRUE))
  
  
  expect_equal(df4,  read.xlsx(wb, sheet = 4))
  expect_equal(df4,  read.xlsx(wb, sheet = "Sheet 4"))
  expect_equal(df4,  read.xlsx(tempFile, sheet = 4))
  expect_equal(df4,  read.xlsx(tempFile, sheet = "Sheet 4"))
  
  expect_equal(df2,  read.xlsx(wb, sheet = 5, rowNames = TRUE))
  expect_equal(df2,  read.xlsx(wb, sheet = "Sheet 2", rowNames = TRUE))
  expect_equal(df2,  read.xlsx(tempFile, sheet = 5, rowNames = TRUE))
  expect_equal(df2,  read.xlsx(tempFile, sheet = "Sheet 2", rowNames = TRUE))
  

  names(wb)

  





  removeWorksheet(wb, sheet = 3)
expect_equal(names(wb),  c("Sheet 1", "Sheet 3", "Sheet 4", "Sheet 2"))
openXL(wb)
  removeWorksheet(wb, sheet = "Sheet 2")

wb$sheetOrder


  
#   removeWorksheet(wb, sheet = 1)            
               
# 
  openXL(wb)
  

  
  
  
  
#   
#   
#   
#   unlink(tempFile, recursive = TRUE, force = TRUE)
#   rm(wb)
#   
#   
# })
# 

