



context("Named Regions")



test_that("Maintaining Named Regions on Load", {
 

  ## create named regions
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  addWorksheet(wb, "Sheet 2")
  
  ## specify region
  writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
  createNamedRegion(wb = wb,
                    sheet = 1,
                    name = "iris",
                    rows = 1:(nrow(iris)+1),
                    cols = 1:ncol(iris))
  
  
  ## using writeData 'name' argument
  writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
  

  ## Named region size 1
  writeData(wb, sheet = 2, x = 99, name = "region1", startCol = 3, startRow = 3)

  ## save file for testing
  out_file <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, out_file, overwrite = TRUE)
  
  
  expect_equal(object = getNamedRegions(wb), expected = getNamedRegions(out_file))
  
  df1 <- read.xlsx(wb, namedRegion = "iris")
  df2 <- read.xlsx(out_file, namedRegion = "iris")
  expect_equal(object = df1, expected = df2)
  
  df1 <- read.xlsx(wb, namedRegion = "region1")
  expect_equal(object = class(df1), expected = "data.frame")
  expect_equal(object = nrow(df1), expected = 0)
  expect_equal(object = ncol(df1), expected = 1)
  
  df1 <- read.xlsx(wb, namedRegion = "region1", colNames = FALSE)
  expect_equal(object = class(df1), expected = "data.frame")
  expect_equal(object = nrow(df1), expected = 1)
  expect_equal(object = ncol(df1), expected = 1)
  
  df1 <- read.xlsx(wb, namedRegion = "region1", rowNames = TRUE)
  expect_equal(object = class(df1), expected = "data.frame")
  expect_equal(object = nrow(df1), expected = 0)
  expect_equal(object = ncol(df1), expected = 0)
  
  
})








