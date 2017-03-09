



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

test_that("Correctly Loading Named Regions Created in Excel",{
  
  # Load an excel workbook (in the repo, it's located in the /inst folder;
  # when installed on the user's system, it is located in the installation folder
  # of the package)
  filename <- system.file("namedRegions.xlsx", package = "openxlsx")
  
  # Load this workbook. We will test read.xlsx by passing both the object wb and
  # the filename. Both should produce the same results.
  wb <- loadWorkbook(filename)
  
  # NamedTable refers to Sheet1!$C$5:$D$8
  table_f <- read.xlsx(filename,
                       namedRegion = "NamedTable")
  table_w <- read.xlsx(wb,
                       namedRegion = "NamedTable")
  
  expect_equal(object = table_f, expected = table_w)
  expect_equal(object = class(table_f), expected = "data.frame")
  expect_equal(object = ncol(table_f), expected = 2)
  expect_equal(object = nrow(table_f), expected = 3)
  
  # NamedCell refers to Sheet1!$C$2
  # This proeduced an error in an earlier version of the pacage when the object
  # wb was passed, but worked correctly when the filename was passed to read.xlsx
  cell_f <- read.xlsx(filename,
                       namedRegion = "NamedCell",
                       colNames = FALSE,
                       rowNames = FALSE)
  
  cell_w <- read.xlsx(wb,
                       namedRegion = "NamedCell",
                       colNames = FALSE,
                       rowNames = FALSE)
  
  expect_equal(object = cell_f, expected = cell_w)
  expect_equal(object = class(cell_f), expected = "data.frame")
  expect_equal(object = ncol(cell_f), expected = 1)
  expect_equal(object = nrow(cell_f), expected = 1)
  
  # NamedCell2 refers to Sheet1!$C$2:$C$2
  cell2_f <- read.xlsx(filename,
                       namedRegion = "NamedCell2",
                       colNames = FALSE,
                       rowNames = FALSE)
  
  cell2_w <- read.xlsx(wb,
                       namedRegion = "NamedCell2",
                       colNames = FALSE,
                       rowNames = FALSE)
  
  expect_equal(object = cell2_f, expected = cell2_w)
  expect_equal(object = class(cell2_f), expected = "data.frame")
  expect_equal(object = ncol(cell2_f), expected = 1)
  expect_equal(object = nrow(cell2_f), expected = 1)
  
})







