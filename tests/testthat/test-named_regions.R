



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




test_that("Missing rows in named regions", {

  temp_file <- tempfile(fileext = ".xlsx")
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  ## create region
  writeData(wb, sheet = 1, x = iris[1:11,], startCol = 1, startRow = 1)
  deleteData(wb, sheet = 1, col = 1:2, rows = c(6, 6))
  
  createNamedRegion(wb = wb,
                    sheet = 1,
                    name = "iris",
                    rows = 1:(5+1),
                    cols = 1:2)
  createNamedRegion(wb = wb,
                    sheet = 1,
                    name = "iris2",
                    rows = 1:(5+2),
                    cols = 1:2)
  
  ## iris region is rows 1:6 & cols 1:2
  ## iris2 region is rows 1:7 & cols 1:2
  
  ## row 6 columns 1 & 2 are blank
  expect_equal(getNamedRegions(wb)[1:2], c("iris", "iris2"), ignore.attributes = TRUE)
  expect_equal(attr(getNamedRegions(wb), "sheet"), c("Sheet 1", "Sheet 1"))
  expect_equal(attr(getNamedRegions(wb), "position"), c("A1:B6", "A1:B7"))
  
  ######################################################################## from Workbook
  
  ## Skip empty rows
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(4, 2))
  
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(5, 2))
  
  
  ## Keep empty rows
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(5, 2))
  
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(6, 2))
  
  
  
  ######################################################################## from file
  saveWorkbook(wb, file = temp_file, overwrite = TRUE)
  
  ## Skip empty rows
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(4, 2))
  
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = TRUE)
  expect_equal(dim(x), c(5, 2))
  
  
  ## Keep empty rows
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(5, 2))
  
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyRows = FALSE)
  expect_equal(dim(x), c(6, 2))
  
  unlink(temp_file)
  
})





test_that("Missing columns in named regions", {
  
  temp_file <- tempfile(fileext = ".xlsx")
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  ## create region
  writeData(wb, sheet = 1, x = iris[1:11,], startCol = 1, startRow = 1)
  deleteData(wb, sheet = 1, col = 2, rows = 1:12, gridExpand = TRUE)

  createNamedRegion(wb = wb,
                    sheet = 1,
                    name = "iris",
                    rows = 1:5,
                    cols = 1:2)
  
  createNamedRegion(wb = wb,
                    sheet = 1,
                    name = "iris2",
                    rows = 1:5,
                    cols = 1:3)
  
  ## iris region is rows 1:5 & cols 1:2
  ## iris2 region is rows 1:5 & cols 1:3
  
  ## row 6 columns 1 & 2 are blank
  expect_equal(getNamedRegions(wb)[1:2], c("iris", "iris2"), ignore.attributes = TRUE)
  expect_equal(attr(getNamedRegions(wb), "sheet"), c("Sheet 1", "Sheet 1"))
  expect_equal(attr(getNamedRegions(wb), "position"), c("A1:B5", "A1:C5"))
  
  ######################################################################## from Workbook
  
  ## Skip empty cols
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 1))
  
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 2))
  
  
  ## Keep empty cols
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 1))
  
  x <- read.xlsx(xlsxFile = wb, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 3))
  
  
  
  ######################################################################## from file
  saveWorkbook(wb, file = temp_file, overwrite = TRUE)
  
  ## Skip empty cols
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 1))
  
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = TRUE)
  expect_equal(dim(x), c(4, 2))
  
  
  ## Keep empty cols
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 1))
  
  x <- read.xlsx(xlsxFile = temp_file, namedRegion = "iris2", colNames = TRUE, skipEmptyCols = FALSE)
  expect_equal(dim(x), c(4, 3))
  
  unlink(temp_file)
  
})
