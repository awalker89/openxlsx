



context("Skip Empty Rows/Cols")





test_that("skip empty cols 1", {
  
  
  xlsxfile <- tempfile()
  x <- data.frame("a" = c(1, NA, NA, 2), "b" = c(1, NA, NA, 3))
  y <- data.frame("x" = c(1, NA, NA, 2), "y" = c(1, NA, NA, 3))
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  writeData(wb, sheet = 1, x = x)
  writeData(wb, sheet = 1, x = y, startCol = 4)
  
  saveWorkbook(wb, file = xlsxfile)
  
  
  ## from file
  res <- readWorkbook(xlsxfile, skipEmptyRows = FALSE, skipEmptyCols = FALSE)
  expect_equal(ncol(res), 5)
  expect_equal(nrow(res), 4)
  
  ## from file
  res <- readWorkbook(xlsxfile, skipEmptyRows = TRUE, skipEmptyCols = TRUE)
  expect_equal(ncol(res), 4)
  expect_equal(nrow(res), 2)
  expect_equivalent(cbind(x, y)[c(1, 4), ], res)
  
  ## from file
  res <- readWorkbook(xlsxfile, skipEmptyRows = FALSE, skipEmptyCols = TRUE)
  expect_equal(ncol(res), 4)
  expect_equal(nrow(res), 4)
  expect_equivalent(cbind(x, y), res)
  
  ## from file
  res <- readWorkbook(xlsxfile, skipEmptyRows = TRUE, skipEmptyCols = FALSE)
  expect_equal(ncol(res), 5)
  expect_equal(nrow(res), 2)
  expect_true(all(is.na(res$X3)))
  
  
  
  
  #############################################################################
  ## Workbook object
  
  ## Workbook object
  wb <- loadWorkbook(xlsxfile)
  
  ## from workbook object
  res <- readWorkbook(wb, skipEmptyRows = FALSE, skipEmptyCols = FALSE)
  expect_equal(ncol(res), 5)
  expect_equal(nrow(res), 4)
  
  ## from workbook object
  res <- readWorkbook(wb, skipEmptyRows = TRUE, skipEmptyCols = TRUE)
  expect_equal(ncol(res), 4)
  expect_equal(nrow(res), 2)
  expect_equivalent(cbind(x, y)[c(1, 4), ], res)
  
  ## from workbook object
  res <- readWorkbook(wb, skipEmptyRows = FALSE, skipEmptyCols = TRUE)
  expect_equal(ncol(res), 4)
  expect_equal(nrow(res), 4)
  expect_equivalent(cbind(x, y), res)
  
  ## from workbook object
  res <- readWorkbook(wb, skipEmptyRows = TRUE, skipEmptyCols = FALSE)
  expect_equal(ncol(res), 5)
  expect_equal(nrow(res), 2)
  expect_true(all(is.na(res$X3)))
  
  
  
})





test_that("skip empty cols numeric only", {
  
  ## Numeric types
  df1 <- data.frame("x1" = 1:3, "x2" = 5:7)
  df2 <- data.frame("x3" = 1:3, "x4" = 5:7)
  df3 <- data.frame("x5" = 1:3, "x6" = 5:7)
  df4 <- data.frame("x7" = 1:3, "x8" = 5:7)
  
  tmp_file <- tempfile(fileext = ".xlsx")
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  writeData(wb, sheet = 1, x = df1, startCol = 1)
  writeData(wb, sheet = 1, x = df2, startCol = 4)
  writeData(wb, sheet = 1, x = df3, startCol = 7)
  writeData(wb, sheet = 1, x = df4, startCol = 10)
  
  saveWorkbook(wb, file = tmp_file, overwrite = TRUE)
  
  x <- read.xlsx(tmp_file, sheet = 1, skipEmptyCols = FALSE)
  y <- read.xlsx(loadWorkbook(tmp_file), sheet = 1, skipEmptyCols = FALSE)
  
  expect_equal(dim(x), c(3, 11))
  expect_equal(dim(y), c(3, 11))
  expect_equal(x, y)
  
  x <- read.xlsx(tmp_file, sheet = 1, skipEmptyCols = TRUE)
  y <- read.xlsx(loadWorkbook(tmp_file), sheet = 1, skipEmptyCols = TRUE)
  
  expect_equal(dim(x), c(3, 8))
  expect_equal(dim(y), c(3, 8))
  expect_equal(x, y)
  
  unlink(tmp_file, force = TRUE)
  
})





test_that("skip empty cols mixed types", {
  
  df1 <- data.frame("x1" = 1:3, "x2" = c("A", "B", "C"))
  df2 <- data.frame("x3" = 1:3, "x4" = c("A", "B", "C"))
  df3 <- data.frame("x5" = 1:3, "x6" = c("A", "B", "C"))
  df4 <- data.frame("x7" = 1:3, "x8" = c("A", "B", "C"))
  
  tmp_file <- tempfile(fileext = ".xlsx")
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  writeData(wb, sheet = 1, x = df1, startCol = 1)
  writeData(wb, sheet = 1, x = df2, startCol = 4)
  writeData(wb, sheet = 1, x = df3, startCol = 7)
  writeData(wb, sheet = 1, x = df4, startCol = 10)
  
  saveWorkbook(wb, file = tmp_file, overwrite = TRUE)
  
  x <- read.xlsx(tmp_file, sheet = 1, skipEmptyCols = FALSE)
  y <- read.xlsx(loadWorkbook(tmp_file), sheet = 1, skipEmptyCols = FALSE)
  
  expect_equal(dim(x), c(3, 11))
  expect_equal(dim(y), c(3, 11))
  expect_equal(x, y)
  
  x <- read.xlsx(tmp_file, sheet = 1, skipEmptyCols = TRUE)
  y <- read.xlsx(loadWorkbook(tmp_file), sheet = 1, skipEmptyCols = TRUE)
  
  expect_equal(dim(x), c(3, 8))
  expect_equal(dim(y), c(3, 8))
  expect_equal(x, y)
  
  unlink(tmp_file, force = TRUE)
  
})





