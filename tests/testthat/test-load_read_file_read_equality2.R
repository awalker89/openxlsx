


context("Reading from workbook is identical to reading from file 2")



test_that("Reading from new workbook 2 ", {
  
  ## data
  genDf <- function(){
    data.frame("Date" = Sys.Date()-0:4,
               "Logical" = c(TRUE, FALSE, TRUE, TRUE, FALSE),
               "Currency" = -2:2,
               "Accounting" = -2:2,
               "hLink" = "http://cran.r-project.org/", 
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

