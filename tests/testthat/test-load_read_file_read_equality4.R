

context("Reading from workbook is identical to reading from file 4")


test_that("Reading from new workbook cols/rows", {
  
  wb <- createWorkbook()
  for(i in 1:4)
    addWorksheet(wb, sprintf('Sheet %s', i))
  
  tempFile <- file.path(tempdir(), "temp.xlsx")
  
  ## 1
  writeData(wb, sheet = 1, x = mtcars, colNames = TRUE, rowNames = FALSE)
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  cols <- 1:3
  rows <- 1:10
  x <- read.xlsx(wb, 1, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  y <- read.xlsx(tempFile, 1, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  
  df <- mtcars[sort((rows-1)[(rows-1) <= nrow(mtcars)]), sort(cols[cols <= ncol(mtcars)])]
  rownames(df) <- 1:nrow(df)
  
  expect_equal(object = x, expected = y)
  expect_equal(object = x, expected = df)
  
  
  
  ## 2
  writeData(wb, sheet = 2, x = mtcars, colNames = TRUE, rowNames = FALSE, startRow = 10, startCol = 5)
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  cols <- 1:300
  rows <- 1:1000
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  y <- read.xlsx(tempFile, sheet = 2, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  
  # 
  expect_equal(object = mtcars, expected = x, check.attributes = FALSE)
  expect_equal(object = x, expected = y, check.attributes = TRUE)
  expect_equal(object = colnames(mtcars), expected = colnames(x), check.attributes = FALSE)
  
  
  cols <- 1:3
  rows <- 12:13
  x <- read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  y <- read.xlsx(tempFile, sheet = 2, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  
  
  expect_equal(object = NULL, expected = x, check.attributes = FALSE)
  expect_equal(object = NULL, expected = y, check.attributes = TRUE)
  
  
  
  ## 3
  writeData(wb, sheet = 3, x = mtcars, colNames = TRUE, rowNames = FALSE)
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  cols <- c(2, 4, 6)
  rows <- seq(1, 31, by = 2)
  
  x <- read.xlsx(wb, sheet = 3, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  y <- read.xlsx(tempFile, sheet = 3, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols)
  
  df <- mtcars[sort((rows-1)[(rows-1) <= nrow(mtcars)]), sort(cols[cols <= ncol(mtcars)])]
  rownames(df) <- 1:nrow(df)
  
  expect_equal(object = x, expected = y, check.attributes = FALSE)
  expect_equal(object = df, expected = x, check.attributes = FALSE)
  
  
  
  ## 4
  writeData(wb, sheet = 4, x = mtcars, colNames = TRUE, rowNames = TRUE)
  saveWorkbook(wb, tempFile, overwrite = TRUE)
  
  cols <- c(1, 6, 12)
  rows <- seq(1, 31, by = 2)
  x <- read.xlsx(wb, sheet = 4, colNames = TRUE, rowNames = TRUE, rows = rows, cols = cols)  
  y <- read.xlsx(tempFile, sheet = 4, colNames = TRUE, rowNames = TRUE, rows = rows, cols = cols)  
  
  df <- mtcars[sort((rows-1)[(rows-1) <= nrow(mtcars)]),cols[-1]-1]
  expect_equal(object = x, expected = y, check.attributes = FALSE)
  expect_equal(object = df, expected = x, check.attributes = FALSE)
  
  rm(wb)
  unlink(tempFile, recursive = TRUE, force = TRUE)
  
})



