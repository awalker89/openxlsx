




context("Read Sources")



test_that("read.xlsx from different sources", {
  
  ## URL
  xlsxFile <- "https://github.com/awalker89/openxlsx/raw/master/inst/readTest.xlsx"
  df_url <- read.xlsx(xlsxFile)
  
  ## File 
  xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
  df_file <- read.xlsx(xlsxFile)
  
  expect_true(all.equal(df_url, df_file), label = "Read from URL")
  
  
  ## Non-existing URL
  xlsxFile <- "https://github.com/awalker89/openxlsx/raw/master/inst/readTest2.xlsx"
  expect_error(read.xlsx(xlsxFile), regexp = "cannot open URL")
  
  
  ## Non-existing File
  xlsxFile <- file.path(dirname(system.file("readTest.xlsx", package = "openxlsx")), "readTest00.xlsx")
  expect_error(read.xlsx(xlsxFile), regexp = "File does not exist.")
  
  
})




test_that("loadWorkbook from different sources", {
  
  ## URL
  xlsxFile <- "https://github.com/awalker89/openxlsx/raw/master/inst/readTest.xlsx"
  wb_url <- loadWorkbook(xlsxFile)
  
  ## File 
  xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
  wb_file <- loadWorkbook(xlsxFile)
  
  ## check
  expect_true(all.equal.Workbook(wb_url, wb_file), "Loading from URL vs local not equal")
  
  
  
  
})



test_that("getDateOrigin from different sources", {
  
  ## URL
  xlsxFile <- "https://github.com/awalker89/openxlsx/raw/master/inst/readTest.xlsx"
  origin_url <- getDateOrigin(xlsxFile)
  
  ## File 
  xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
  origin_file <- getDateOrigin(xlsxFile)
  
  ## check
  expect_equal(origin_url, origin_file)
  expect_equal(origin_url, "1900-01-01")

  
})





