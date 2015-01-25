






context("Removing worksheets.")

test_that("Deleting worksheets", {
  
  
  tempFile <- file.path(tempdir(), "temp.xlsx")
  genWS <- function(wb, sheetName){
    addWorksheet(wb, sheetName)
    writeDataTable(wb, sheetName, data.frame("X" = sprintf("This is sheet: %s", sheetName)), colNames = FALSE)
  }
  
  wb <- createWorkbook()
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  genWS(wb, "Sheet 3")
    
  expect_equal(names(wb), c("Sheet 1", "Sheet 2", "Sheet 3"))
  
  removeWorksheet(wb, sheet = 1)
  expect_equal(names(wb), c("Sheet 2", "Sheet 3"))
  
  
  removeWorksheet(wb, sheet = 1)
  expect_equal(names(wb), c("Sheet 3"))  
  
  
  ## add to end
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  expect_equal(names(wb), c("Sheet 3", "Sheet 1", "Sheet 2"))  
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  ## re-load & re-order worksheets 
  
  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb), c("Sheet 3", "Sheet 1", "Sheet 2"))  
  
  writeData(wb, sheet = "Sheet 2", x = iris[1:10, 1:4], startRow = 5)
  expect_equal(iris[1:10, 1:4], read.xlsx(wb, "Sheet 2", startRow = 5))  

  
  writeData(wb, sheet = 1, x = iris[1:20, 1:4], startRow = 5)
  expect_equal(iris[1:20, 1:4], read.xlsx(wb, "Sheet 3", startRow = 5))  
  
  
  removeWorksheet(wb, sheet = 1)
  expect_equal("This is sheet: Sheet 1", read.xlsx(wb, 1, startRow = 1)[[1]])  
  
  removeWorksheet(wb, sheet = 2)
  expect_equal("This is sheet: Sheet 1", read.xlsx(wb, 1, startRow = 1)[[1]])  
  
  removeWorksheet(wb, sheet = 1)
  expect_equal(names(wb), character(0)) 
  
    
    
  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
  
  
})


