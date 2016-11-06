




context("Read Sources")



test_that("Read from different sources", {
  
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




