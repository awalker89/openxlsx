


context("Reading from workbook is identical to reading from file readTest.xlsx")



test_that("Reading example workbook", {
  
  xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
  wb <- loadWorkbook(xlsxFile)
  
  ## sheet 1
  sheet <- 1
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(x, y)
  
  
  
  
  ## sheet 2
  sheet <- 2
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(x, y)
  
  
  
  ## sheet 3
  sheet <- 3
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- suppressWarnings(read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE))
  y <- suppressWarnings(read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE))
  expect_equal(x, y)
  
  x <- suppressWarnings(read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE))
  y <- suppressWarnings(read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE))
  expect_equal(x, y)
  
  x <- suppressWarnings(read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE))
  y <- suppressWarnings(read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE))
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(x, y)
  
  
  
  ## sheet 4
  sheet <- 4
  x <- read.xlsx(xlsxFile, sheet)
  y <- read.xlsx(wb, sheet)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, startRow = 3, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = TRUE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = TRUE)
  expect_equal(x, y)
  
  x <- read.xlsx(xlsxFile, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  y <- read.xlsx(wb, sheet, rows = 2:4, colNames = FALSE, detectDates = FALSE)
  expect_equal(x, y)
  
  
  x <- read.xlsx(xlsxFile, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  y <- read.xlsx(wb, sheet, colNames = FALSE, detectDates = FALSE, cols = 2:4)
  expect_equal(x, y)
  
  
})

