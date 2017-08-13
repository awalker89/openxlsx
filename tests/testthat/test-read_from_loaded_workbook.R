


context("Reading from workbook is identical to reading from file readTest.xlsx")



test_that("Reading example workbook", {
  
  xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
  wb <- loadWorkbook(xlsxFile)
  
  ## sheet 1
  sheet <- 1
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(dim(x), c(10, 7))
  expect_equal(dim(y), c(10, 7))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, detectDates = TRUE)
  expect_equal(dim(x), c(10, 7))
  expect_equal(dim(y), c(10, 7))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(9, 5))
  expect_equal(dim(y), c(9, 5))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  expect_equal(dim(x), c(2, 6))
  expect_equal(dim(y), c(2, 6))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(3, 6))
  expect_equal(dim(y), c(3, 6))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  expect_equal(dim(x), c(3, 6))
  expect_equal(dim(y), c(3, 6))
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(dim(x), c(9, 2))
  expect_equal(dim(y), c(9, 2))
  expect_equal(x, y)
  
  
  
  
  ## sheet 2
  sheet <- 2
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(dim(x), c(33, 9))
  expect_equal(dim(y), c(33, 9))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(32, 9))
  expect_equal(dim(y), c(32, 9))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  expect_equal(dim(x), c(2, 9))
  expect_equal(dim(y), c(2, 9))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(3, 9))
  expect_equal(dim(y), c(3, 9))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  expect_equal(dim(x), c(3, 9))
  expect_equal(dim(y), c(3, 9))
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(dim(x), c(21, 3))
  expect_equal(dim(y), c(21, 3))
  expect_equal(x, y)
  
  
  
  ## sheet 3
  sheet <- 3
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(dim(x), c(2083, 5))
  expect_equal(dim(y), c(2083, 5))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(2084, 5))
  expect_equal(dim(y), c(2084, 5))
  expect_equal(x, y)
  
  x <- suppressWarnings(read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE))
  y <- suppressWarnings(read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE))
  expect_equal(dim(x), NULL)
  expect_equal(dim(y), NULL)
  expect_equal(x, y)
  
  x <- suppressWarnings(read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE))
  y <- suppressWarnings(read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE))
  expect_equal(dim(x), NULL)
  expect_equal(dim(y), NULL)
  expect_equal(x, NULL)
  expect_equal(y, NULL)
  expect_equal(x, y)
  
  x <- suppressWarnings(read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE))
  y <- suppressWarnings(read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE))
  expect_equal(dim(x), NULL)
  expect_equal(dim(y), NULL)
  expect_equal(x, NULL)
  expect_equal(y, NULL)
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(dim(x), c(2084, 2))
  expect_equal(dim(y), c(2084, 2))
  expect_equal(x, y)
  
  
  
  ## sheet 5
  sheet <- 5
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(dim(x), c(271, 297))
  expect_equal(dim(y), c(271, 297))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(270, 297))
  expect_equal(dim(y), c(270, 297))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  expect_equal(dim(x), c(2, 297))
  expect_equal(dim(y), c(2, 297))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  expect_equal(dim(x), c(3, 297))
  expect_equal(dim(y), c(3, 297))
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  expect_equal(dim(x), c(3, 297))
  expect_equal(dim(y), c(3, 297))
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(dim(x), c(272, 3))
  expect_equal(dim(y), c(272, 3))
  expect_equal(x, y)
  
  
})






test_that("Load read - Skip Empty rows/cols", {
  
  xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
  wb <- loadWorkbook(xlsxFile)
  
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = TRUE, colNames = FALSE)
  expect_equal(nrow(x), 5L)
  expect_equal(ncol(x), 4L)
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = TRUE, colNames = TRUE)
  expect_equal(nrow(x), 5L - 1L)
  expect_equal(ncol(x), 4L)
  
  
  ##############################################################
  ## FALSE FALSE FALSE
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = FALSE, colNames = FALSE)
  expect_equal(nrow(x), 6L)
  expect_equal(ncol(x), 8L)
  
  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))
  
  ## NA rows
  expect_true(all(is.na(x[3,])))
  
  
  
  ##############################################################
  ## FALSE FALSE TRUE
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = FALSE, colNames = TRUE)
  expect_equal(nrow(x), 6L - 1L)
  expect_equal(ncol(x), 8L)
  
  
  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))
  
  ## NA rows
  expect_true(all(is.na(x[2,])))
  
  
  
  ##############################################################
  ## FALSE TRUE FALSE
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = TRUE, colNames = FALSE)
  expect_equal(nrow(x), 5L)
  expect_equal(ncol(x), 8L)
  
  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))
  
  
  
  
  ##############################################################
  ## FALSE TRUE TRUE
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = TRUE, colNames = TRUE)
  expect_equal(nrow(x), 5L - 1L)
  expect_equal(ncol(x), 8L)
  
  
  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))
  
  
  
  ##############################################################
  ## TRUE FALSE FALSE
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = FALSE, colNames = FALSE)
  expect_equal(nrow(x), 6L)
  expect_equal(ncol(x), 4L)
  
  ## NA rows
  expect_true(all(is.na(x[3,])))
  
  
  
  ##############################################################
  ## TRUE FALSE TRUE
  
  x <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = FALSE, colNames = TRUE)
  expect_equal(nrow(x), 6L - 1L)
  expect_equal(ncol(x), 4L)
  
  
  ## NA rows
  expect_true(all(is.na(x[2,])))
  
  
  
})





