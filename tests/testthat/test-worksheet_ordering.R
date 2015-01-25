






context("Re-ordering worksheets.")



test_that("Worksheet ordering from new Workbook", {
  
  
  genWS <- function(wb, sheetName){
    addWorksheet(wb, sheetName)
    writeDataTable(wb, sheetName, data.frame("X" = sprintf("This is sheet: %s", sheetName)), colNames = FALSE)
  }
  
  wb <- createWorkbook()
  genWS(wb, "Sheet 1")
  genWS(wb, "Sheet 2")
  genWS(wb, "Sheet 3")
  
  
  tempFile <- file.path(tempdir(), "orderingTest.xlsx")
  #   tempFile <- file.path("c:/users/alex/desktop", "orderingTest.xlsx")
  
  file <- xlsxFile <- tempFile
  
  ## no ordering
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)
  expect_equal(names(wb),  sprintf("Sheet %s", 1:3))
  
  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb),  sprintf("Sheet %s", 1:3))
  
  
  ## re-order doesnt do anything
  worksheetOrder(wb) <- c(3, 2, 1)
  expect_equal(names(wb),  sprintf("Sheet %s", 1:3))
  
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)
  expect_equal(names(wb),  sprintf("Sheet %s", 1:3))
  
  
  
  ## reloading - reordered
  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb),  sprintf("Sheet %s", 3:1))
  
  x <- read.xlsx(tempFile, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(tempFile, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(tempFile, sheet = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  
  ## reloading  - reordered - reading from the workbook object  
  x <- read.xlsx(wb, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(wb, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(wb, sheet  = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  
  
  ## save and re-load again  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb),  sprintf("Sheet %s", 3:1))
  
  x <- read.xlsx(wb, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(wb, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(wb, sheet = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  x <- read.xlsx(wb, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(wb, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(wb, sheet  = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  
  
  
  ###### re-order again
  worksheetOrder(wb) <- c(2, 3, 1)
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  x <- read.xlsx(tempFile, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(tempFile, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  x <- read.xlsx(tempFile, sheet = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  
  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb),  sprintf("Sheet %s", c(2, 1, 3)))
  
  x <- read.xlsx(wb, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(wb, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  x <- read.xlsx(wb, sheet = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  
  
  
  ## add a worksheet
  genWS(wb, sheetName = "Sheet 4")
  
  x <- read.xlsx(wb, sheet = 4)[[1]]
  expect_equal(x, "This is sheet: Sheet 4")
  
  ## re-order and add worksheet then save
  worksheetOrder(wb) <- c(3, 1, 4, 2)  
  names(wb)
  
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  ## read from file
  x <- read.xlsx(tempFile, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(tempFile, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(tempFile, sheet = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 4")
  
  x <- read.xlsx(tempFile, sheet = 4)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  x <- read.xlsx(tempFile, sheet = "Sheet 3")[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(tempFile, sheet = "Sheet 2")[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(tempFile, sheet = "Sheet 4")[[1]]
  expect_equal(x, "This is sheet: Sheet 4")
  
  x <- read.xlsx(tempFile, sheet = "Sheet 1")[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  
  
  
  
  
  
  ## read from workbook
  wb <- loadWorkbook(tempFile)
  x <- read.xlsx(wb, sheet = 1)[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(wb, sheet = 2)[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(wb, sheet = 3)[[1]]
  expect_equal(x, "This is sheet: Sheet 4")
  
  x <- read.xlsx(wb, sheet = 4)[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  
  
  
  ## read from workbook using name
  wb <- loadWorkbook(tempFile)
  x <- read.xlsx(wb, sheet = "Sheet 3")[[1]]
  expect_equal(x, "This is sheet: Sheet 3")
  
  x <- read.xlsx(wb, sheet = "Sheet 2")[[1]]
  expect_equal(x, "This is sheet: Sheet 2")
  
  x <- read.xlsx(wb, sheet = "Sheet 1")[[1]]
  expect_equal(x, "This is sheet: Sheet 1")
  
  x <- read.xlsx(wb, sheet = "Sheet 4")[[1]]
  expect_equal(x, "This is sheet: Sheet 4")
  
    
  writeData(wb, sheet = "Sheet 3", iris[1:10, 1:4], startRow = 5)
  x  <- read.xlsx(wb, sheet = "Sheet 3", startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:10, 1:4])
  

  writeData(wb, sheet = 4, iris[1:20, 1:4], startRow = 5)
  x  <- read.xlsx(wb, sheet = 4, startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:20, 1:4])
  
  
  writeData(wb, sheet = 2, iris[1:30, 1:4], startRow = 5)
  x  <- read.xlsx(wb, sheet = 2, startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:30, 1:4])
  
  
  ## reading from saved file
  saveWorkbook(wb, tempFile, TRUE)
  
  x  <- read.xlsx(tempFile, sheet = "Sheet 3", startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:10, 1:4])
  
  x  <- read.xlsx(tempFile, sheet = 4, startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:20, 1:4])
  
  x  <- read.xlsx(tempFile, sheet = 2, startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:30, 1:4])
  
  
  ## And finally load again
  wb <- loadWorkbook(tempFile)
  
  x  <- read.xlsx(wb, sheet = "Sheet 3", startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:10, 1:4])

  x  <- read.xlsx(wb, sheet = 4, startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:20, 1:4])
  
  x  <- read.xlsx(wb, sheet = 2, startRow = 5, colNames = TRUE)
  expect_equal(x, iris[1:30, 1:4])
  
  
  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
  
})



test_that("Worksheet ordering from new Workbook", {
  
  
  tempFile <- file.path(tempdir(), "temp.xlsx")
  
  wb <- createWorkbook()
  addWorksheet(wb = wb, sheetName = "Sheet 1", gridLines = FALSE)
  writeDataTable(wb = wb, sheet = 1, x = iris)
  
  addWorksheet(wb = wb, sheetName = "mtcars (Sheet 2)", gridLines = FALSE)
  writeData(wb = wb, sheet = 2, x = mtcars)
  
  addWorksheet(wb = wb, sheetName = "Sheet 3", gridLines = FALSE)
  writeData(wb = wb, sheet = 3, x = Formaldehyde)
  
  worksheetOrder(wb)
  names(wb)
  worksheetOrder(wb) <- c(1,3,2) # switch position of sheets 2 & 3 
  
  names(wb)
  writeData(wb, 2, 'This is still the "mtcars" worksheet', startCol = 15)
  
  names(wb)
  writeData(wb, "Sheet 3", "writing to sheet 3", startCol = 15)
  
  worksheetOrder(wb)
  names(wb)  ## ordering within workbook is not changed
  
  saveWorkbook(wb, tempFile,  overwrite = TRUE)
  
  worksheetOrder(wb) <- c(3,2,1)
  saveWorkbook(wb, tempFile,  overwrite = TRUE)

  
  wb <- loadWorkbook(tempFile)
  worksheetOrder(wb) <- c(3,2,1)

  
  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)

    
})



