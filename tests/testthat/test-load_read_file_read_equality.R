



context("Reading from workbook is identical to reading from file")



test_that("Reading from new workbook", {
  
  wb <- createWorkbook()
  for(i in 1:4)
    addWorksheet(wb, sprintf('Sheet %s', i))
  
  
  ## colNames = TRUE, rowNames = TRUE
  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = TRUE, startRow = 10, startCol = 5)
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = TRUE)
  
  
  ## colNames = TRUE, rowNames = FALSE
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5)
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), check.attributes = FALSE)
  
  ## colNames = FALSE, rowNames = TRUE
  writeData(wb, sheet = 3, x = mtcars, colNames = FALSE, rowNames = TRUE, startRow = 2, startCol = 2)
  x <- read.xlsx(wb, sheet = 3, colNames = FALSE, rowNames = TRUE)
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = rownames(mtcars), expected = rownames(x))
  
  
  ## colNames = FALSE, rowNames = FALSE
  writeData(wb, sheet = 4, x = mtcars, colNames = FALSE, rowNames = FALSE, startRow = 12, startCol = 1)
  x <- read.xlsx(wb, sheet = 4, colNames = FALSE, rowNames = FALSE)  
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  
  rm(wb)
  
})











test_that("Empty workbook", {
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  expect_equal(NULL, read.xlsx(wb))  
  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = FALSE))  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = TRUE))
  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = FALSE))
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = FALSE))
  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE))
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE))
  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10))
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10))
  
  
  ## 1 element
  writeData(wb, 1, "a")
  
  x <- read.xlsx(wb)
  expect_equal(nrow(x), 0)  
  expect_equal(names(x), "a")  
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE)
  expect_equal(data.frame("X1" = "a", stringsAsFactors = FALSE), x)  
  
  x <- read.xlsx(wb, sheet = 1, colNames = TRUE)
  expect_equal(nrow(x), 0)  
  expect_equal(names(x), "a") 
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = FALSE)
  expect_equal(data.frame("X1" = "a", stringsAsFactors = FALSE), x)  
  
  x <- read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(nrow(x), 0)  
  expect_equal(names(x), "a") 
  
  
  writeData(wb, 1, Sys.Date())
  x <- read.xlsx(wb)
  expect_equal(nrow(x), 1)  
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE)
  expect_equal(nrow(x), 2)  
  
  x <- read.xlsx(wb, sheet = 1, colNames = TRUE)
  expect_equal(nrow(x), 1)  
  expect_equal(names(x), "x") 
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE)
  expect_equal(class(x[[1]]), "character")
  
  x <- read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = TRUE)
  expect_equal(x[[1]], Sys.Date())
  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10))
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10))
  
  
  addWorksheet(wb, "Sheet 2")
  removeWorksheet(wb, 1)
  
  
  ## 1 date  
  writeData(wb, 1, Sys.Date(), colNames = FALSE)
  
  x <- read.xlsx(wb)
  expect_equal(convertToDate(names(x)), Sys.Date())  
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE)
  x1 <- convertToDate(x[[1]])
  expect_equal(x1, Sys.Date())  
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE)
  expect_equal(class(x[[1]]), "Date")
  expect_equal(x[[1]], Sys.Date())
  
  x <- read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = TRUE)
  expect_equal(convertToDate(names(x)), Sys.Date())
  
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10))
  expect_equal(NULL, read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10))
  
  
  
})




test_that("NAs and NaN values", {
  
  fileName <- file.path(tempdir(), "NaN.xlsx")

  ## data
  a <- data.frame("X" = c(-pi/0, NA, NaN),
                  "Y" = letters[1:3], 
                  "Z" = c(pi/0, 99, NaN),
                  "Z2" = c(1, NaN, NaN),
                  stringsAsFactors = FALSE)
  
  b <- a
  b[b == -Inf] <- NaN
  b[b == Inf] <- NaN
  
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  writeData(wb, 1, a, keepNA = FALSE)
  
  addWorksheet(wb, "Sheet 2")
  writeData(wb, 2, a, keepNA = TRUE)
  
  saveWorkbook(wb, file = fileName, overwrite = TRUE)
  
  ## keepNA = FALSE
  expect_equal(b, read.xlsx(wb))  
  expect_equal(b, read.xlsx(fileName)) 
  

  ## keepNA = TRUE
  expect_equal(b, read.xlsx(wb, sheet = 2))  
  expect_equal(b, read.xlsx(fileName, sheet = 2)) 
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
})
















