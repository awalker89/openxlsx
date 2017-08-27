

context("Writing and reading returns similar objects")

test_that("Writing then reading returns identical data.frame 1", {
  
  curr_wd <- getwd()
  
  ## data
  genDf <- function(){
    
    set.seed(1)
    data.frame("Date" = Sys.Date()-0:4,
               "Logical" = c(TRUE, FALSE, TRUE, TRUE, FALSE),
               "Currency" = -2:2,
               "Accounting" = -2:2,
               "hLink" = "https://CRAN.R-project.org/", 
               "Percentage" = seq(-1, 1, length.out=5),
               "TinyNumber" = runif(5) / 1E9, stringsAsFactors = FALSE)
  }
  
  df <- genDf()
  df
  
  class(df$Currency) <- "currency"
  class(df$Accounting) <- "accounting"
  class(df$hLink) <- "hyperlink"
  class(df$Percentage) <- "percentage"
  class(df$TinyNumber) <- "scientific"
  
  options("openxlsx.dateFormat" = "yyyy-mm-dd")
  
  fileName <- file.path(tempdir(), "allClasses.xlsx")
  write.xlsx(df, file = fileName, overwrite = TRUE)
  
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE)
  expect_equal(object = x, expected = genDf(), check.attributes = FALSE)
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
  expect_equal(object = getwd(), curr_wd)
  
  
  
})




test_that("Writing then reading returns identical data.frame 2", {
  
  curr_wd <- getwd()
  
  ## data.frame of dates
  dates <- data.frame("d1" = Sys.Date() - 0:500)
  for(i in 1:3) dates <- cbind(dates, dates)
  names(dates) <- paste0("d", 1:8)
  
  ## Date Formatting
  wb <- createWorkbook()
  addWorksheet(wb, "Date Formatting", gridLines = FALSE)
  writeData(wb, 1, dates) ## write without styling
  
  ## set default date format
  options("openxlsx.dateFormat" = "yyyy/mm/dd")
  
  ## numFmt == "DATE" will use the date format specified by the above
  addStyle(wb, 1, style = createStyle(numFmt = "DATE"), rows = 2:11, cols = 1, gridExpand = TRUE) 
  
  ## some custom date format examples
  sty <- createStyle(numFmt = "yyyy/mm/dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 2, gridExpand = TRUE)
  
  sty <- createStyle(numFmt = "yyyy/mmm/dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 3, gridExpand = TRUE)
  
  sty <- createStyle(numFmt = "yy / mmmm / dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 4, gridExpand = TRUE)
  
  sty <- createStyle(numFmt = "ddddd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 5, gridExpand = TRUE)
  
  sty <- createStyle(numFmt = "yyyy-mmm-dd")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 6, gridExpand = TRUE)
  
  sty <- createStyle(numFmt = "mm/ dd yyyy")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 7, gridExpand = TRUE)
  
  sty <- createStyle(numFmt = "mm/dd/yy")
  addStyle(wb, 1, style = sty, rows = 2:11, cols = 8, gridExpand = TRUE)
  
  setColWidths(wb, 1, cols = 1:10, widths = 23)
  
  
  fileName <- file.path(tempdir(), "DateFormatting.xlsx")
  write.xlsx(dates, file = fileName, overwrite = TRUE)
  
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE)  
  expect_equal(object = x, expected = dates, check.attributes = FALSE)
  
  
  xNoDateDetection <- read.xlsx(xlsxFile = fileName, detectDates = FALSE)  
  dateOrigin <- getDateOrigin(fileName)
  
  
  expect_equal(object = dateOrigin, expected = "1900-01-01", check.attributes = FALSE)
  
  for(i in 1:ncol(x))
    xNoDateDetection[[i]] <- convertToDate(xNoDateDetection[[i]], origin = dateOrigin)
  
  expect_equal(object = xNoDateDetection, expected = dates, check.attributes = FALSE)
  
  expect_equal(object = getwd(), curr_wd)
  unlink(fileName, recursive = TRUE, force = TRUE)
  
  
})







test_that("Writing then reading rowNames, colNames combinations", {
  
  fileName <- file.path(tempdir(), "tmp.xlsx")
  curr_wd <- getwd()
  
  ## rowNames = colNames = TRUE
  write.xlsx(mtcars, file = fileName, overwrite = TRUE, row.names = TRUE)
  x <- read.xlsx(fileName, sheet = 1, rowNames = TRUE)
  expect_equal(object = x, expected = mtcars, check.attributes = TRUE)
  
  
  ## rowNames = TRUE, colNames = FALSE
  write.xlsx(mtcars, file = fileName, overwrite = TRUE, rowNames = TRUE, colNames = FALSE)
  x <- read.xlsx(fileName, sheet = 1, rowNames = TRUE, colNames = FALSE)
  expect_equal(object = x, expected = mtcars, check.attributes = FALSE)
  expect_equal(object = rownames(x), expected = rownames(mtcars))
  
  
  ## rowNames = FALSE, colNames = TRUE
  write.xlsx(mtcars, file = fileName, overwrite = TRUE, rowNames = FALSE, colNames = TRUE)
  x <- read.xlsx(fileName, sheet = 1, rowNames = FALSE, colNames = TRUE)
  expect_equal(object = x, expected = mtcars, check.attributes = FALSE)
  
  ## rowNames = FALSE, colNames = FALSE
  write.xlsx(mtcars, file = fileName, overwrite = TRUE, rowNames = FALSE, colNames = FALSE)
  x <- read.xlsx(fileName, sheet = 1, rowNames = FALSE, colNames = FALSE)
  expect_equal(object = x, expected = mtcars, check.attributes = FALSE)
  
  expect_equal(object = getwd(), curr_wd)
  unlink(fileName, recursive = TRUE, force = TRUE)

  
})







test_that("Writing then reading returns identical data.frame 3", {
  
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
  
  options("openxlsx.dateFormat" = "yyyy-mm-dd")
  
  fileName <- file.path(tempdir(), "allClasses.xlsx")
  write.xlsx(df, file = fileName, overwrite = TRUE)
  
  
  ## rows, cols combinations
  rows <- 1:4
  cols <- c(1, 3, 5)
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols)
  expect_equal(object = x, expected = genDf()[sort((rows-1)[(rows-1) <= nrow(df)]), sort(cols[cols <= ncol(df)])], check.attributes = FALSE)
  
  
  rows <- 1:4
  cols <- 1:9
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols)
  expect_equal(object = x, expected = genDf()[sort((rows-1)[(rows-1) <= nrow(df)]), sort(cols[cols <= ncol(df)])], check.attributes = FALSE)
  
  
  rows <- 1:200
  cols <- c(5, 99, 2)
  x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols)
  expect_equal(object = x, expected = genDf()[sort((rows-1)[(rows-1) <= nrow(df)]), sort(cols[cols <= ncol(df)])], check.attributes = FALSE)
  
  
  rows <- 1000:900
  cols <- c(5, 99, 2)
  suppressWarnings(x <- read.xlsx(xlsxFile = fileName, detectDates = TRUE, rows = rows, cols = cols))
  expect_equal(object = x, expected = NULL, check.attributes = FALSE)
  
  unlink(fileName, recursive = TRUE, force = TRUE)
  
})






test_that("Writing then reading returns identical data.frame 4", {
  
  ## data
  df <- head(iris[,1:4])
  df[1,2] <- NA
  df[3,1] <- NA
  df[6, 4] <- NA
  
  
  tf <- tempfile(fileext = ".xlsx")
  write.xlsx(x = df, file = tf, keepNA = TRUE)
  x <- read.xlsx(tf)
  
  expect_equal(object = x, expected = df, check.attributes = TRUE)
  unlink(tf, recursive = TRUE, force = TRUE)
  
  
  tf <- tempfile(fileext = ".xlsx")
  write.xlsx(x = df, file = tf, keepNA = FALSE)
  x <- read.xlsx(tf)
  
  expect_equal(object = x, expected = df, check.attributes = TRUE)
  unlink(tf, recursive = TRUE, force = TRUE)
  
})




test_that("Special characters in sheet names", {
  
  tf <- tempfile(fileext = ".xlsx")
  
  ## data
  sheet_name <- "A & B < D > D"
  
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = sheet_name)
  addWorksheet(wb, sheetName = "test")
  writeData(wb, sheet = 1, x = 1:10)
  saveWorkbook(wb = wb, file = tf, overwrite = TRUE)
  
  expect_equal(getSheetNames(tf)[1], sheet_name)
  expect_equal(getSheetNames(tf)[2], "test")
  
  expect_equal(read.xlsx(tf, colNames = FALSE)[[1]], 1:10)
  
  unlink(tf, recursive = TRUE, force = TRUE)
  
})

