



context("Reading from wb object is identical to reading from file")



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
  
  
  expect_equal(NULL, suppressWarnings(read.xlsx(wb)))  
  
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = FALSE)))
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = TRUE)))
  
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = FALSE)))
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = FALSE)))
  
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE)))
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE)))
  
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10)))
  expect_equal(NULL, suppressWarnings(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10)))
  
  
  
  
  expect_warning(read.xlsx(wb))  
  
  expect_warning(read.xlsx(wb, sheet = 1, colNames = FALSE))
  expect_warning(read.xlsx(wb, sheet = 1, colNames = TRUE))
  
  expect_warning(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = FALSE))
  expect_warning(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = FALSE))
  
  expect_warning(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE))
  expect_warning(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE))
  
  expect_warning(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10))
  expect_warning(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10))
  
  
  
  
  
  
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
  
  writeData(wb, 1, Sys.Date(), startCol = 1, startRow = 1)
  x <- read.xlsx(wb)
  expect_equal(nrow(x), 0)  
  expect_equal(convertToDate(as.integer(names(x)[1])), Sys.Date())
  

  x <- read.xlsx(wb, sheet = 1, colNames = FALSE)
  expect_equal(nrow(x), 1)  
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE)
  expect_equal(class(x[[1]]), "Date")
  
  x <- read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE)
  expect_equal(x[[1]], Sys.Date())
  
  x <- suppressWarnings(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10))
  expect_equal(NULL, x)
  
  x <- suppressWarnings(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10))
  expect_equal(NULL, x)
  
  

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
  expect_equal(as.Date(names(x)), Sys.Date())
  
  x <- suppressWarnings(read.xlsx(wb, sheet = 1, colNames = FALSE, skipEmptyRows = TRUE, detectDates = TRUE, rows = 4:10))
  expect_equal(NULL, x)
  
  x <- suppressWarnings(read.xlsx(wb, sheet = 1, colNames = TRUE, skipEmptyRows = TRUE, detectDates = FALSE, cols = 4:10))
  expect_equal(NULL, x)
  

  
  
})




test_that("Reading NAs and NaN values", {
  
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
  
  ## from file
  expected_df <- structure(list(X = c(NA_real_, NA_real_, NA_real_)
                                , Y = c("a", "b", "c")
                                , Z = c(NA, 99, NA)
                                , Z2 = c(1, NA, NA))
                           , .Names = c("X", "Y", "Z", "Z2")
                           , row.names = c(NA, 3L), class = "data.frame")
  
  expect_equal(read.xlsx(fileName), expected_df)
  
  ## from workbook
  expected_df <- structure(list(X = c(NA_real_, NA_real_, NA_real_)
                                , Y = c("a", "b", "c")
                                , Z = c(NA, 99, NA)
                                , Z2 = c(1, NA, NA))
                           , .Names = c("X", "Y", "Z", "Z2")
                           , row.names = c(NA, 3L), class = "data.frame")
  
  expect_equal(read.xlsx(wb), expected_df)
  
  
   
  ## keepNA = FALSE
  expect_equal(read.xlsx(wb), read.xlsx(fileName))  
  expect_equal(b, read.xlsx(wb))  
  expect_equal(b, read.xlsx(fileName)) 
  
  ## keepNA = TRUE
  
  expect_equal(read.xlsx(wb), expected_df)
  expect_equal(read.xlsx(fileName), expected_df)
  
  expect_equal(b, read.xlsx(wb, sheet = 2))  
  expect_equal(b, read.xlsx(fileName, sheet = 2)) 
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
})





test_that("Reading from new workbook 2 ", {
  
  ## data
  genDf <- function(){
    data.frame("Date" = Sys.Date()-0:4,
               "Logical" = c(TRUE, FALSE, TRUE, TRUE, FALSE),
               "Currency" = -2:2,
               "Accounting" = -2:2,
               "hLink" = "https://CRAN.R-project.org/", 
               "Percentage" = seq(-1, 1, length.out=5),
               "TinyNumber" = runif(5) / 1E9, stringsAsFactors = FALSE)
  }
  
  df <- genDf()
  
  class(df$Currency) <- "currency"
  class(df$Accounting) <- "accounting"
  class(df$hLink) <- "hyperlink"
  class(df$Percentage) <- "percentage"
  class(df$TinyNumber) <- "scientific"
  
  options("openxlsx.dateFormat" = NULL)
  
  fileName <- file.path(tempdir(), "allClasses.xlsx")
  wb <- write.xlsx(df, file = fileName, overwrite = TRUE)
  
  x <- read.xlsx(wb, sheet = 1, detectDates = FALSE)
  x[[1]] <- convertToDate(x[[1]])
  expect_equal(object = x, expected = genDf(), check.attributes = FALSE)
  
  
  x <- read.xlsx(wb, sheet = 1, detectDates = TRUE)
  expect_equal(object = x, expected = genDf(), check.attributes = FALSE)
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
})





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
  x <- suppressWarnings(read.xlsx(wb, sheet = 2, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols))
  y <- suppressWarnings(read.xlsx(tempFile, sheet = 2, colNames = TRUE, rowNames = FALSE, rows = rows, cols = cols))
  
  
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










