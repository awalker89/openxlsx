



context("Skip Empty Rows/Cols")



test_that("skip empty rows", {
  

  xlsxfile <- tempfile()
  df <- data.frame("x" = c(1, NA, NA, 2), "y" = c(1, NA, NA, 3))
  
  write.xlsx(df, xlsxfile)
  
  wb <- loadWorkbook(xlsxfile)
  
  df1 <- readWorkbook(xlsxfile, skipEmptyRows = FALSE)
  df2 <- readWorkbook(wb, skipEmptyRows = FALSE)
  
  expect_equal(df, df1)
  expect_equal(df, df2)
  
  
  v <- c("A1", "B1", "A2", "B2", "A5", "B5")
  expect_equal(.Call('openxlsx_calc_number_rows', PACKAGE = 'openxlsx', x = v, skipEmptyRows = TRUE), 3)
  expect_equal(.Call('openxlsx_calc_number_rows', PACKAGE = 'openxlsx', x = v, skipEmptyRows = FALSE), 5)
  
  
  
  ## DONT SKIP
  
  df1 <- readWorkbook(xlsxfile, skipEmptyRows = TRUE)
  df2 <- readWorkbook(wb, skipEmptyRows = TRUE)
  
  expect_equal(nrow(df1), 2)
  expect_equal(nrow(df2), 2)
  
  expect_equivalent(df[c(1,4), ], df1)
  expect_equivalent(df[c(1,4), ], df2)
  
  
  
  
  
})






test_that("skip empty cols", {
  
  
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





