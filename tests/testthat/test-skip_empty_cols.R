



context("Skip Empty Cols")



test_that("skip empty rows", {
  xlsxfile <- tempfile()
  df <- data.frame("x" = c(1, NA, NA, 2), "y" = c(1, NA, NA, 3))

  write.xlsx(df, xlsxfile)

  wb <- loadWorkbook(xlsxfile)

  df1 <- readWorkbook(xlsxfile, skipEmptyRows = FALSE)
  df2 <- readWorkbook(wb, skipEmptyRows = FALSE)

  expect_equal(df, df1)
  expect_equal(df, df2)


  v <- c("A1", "B1", "A2", "B2", "A5", "B5")
  expect_equal(calc_number_rows(x = v, skipEmptyRows = TRUE), 3)
  expect_equal(calc_number_rows(x = v, skipEmptyRows = FALSE), 5)



  ## DONT SKIP

  df1 <- readWorkbook(xlsxfile, skipEmptyRows = TRUE)
  df2 <- readWorkbook(wb, skipEmptyRows = TRUE)

  expect_equal(nrow(df1), 2)
  expect_equal(nrow(df2), 2)

  expect_equivalent(df[c(1, 4), ], df1)
  expect_equivalent(df[c(1, 4), ], df2)
})







test_that("Version 4 fixes from File", {
  fl <- system.file("readTest.xlsx", package = "openxlsx")


  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = TRUE, colNames = FALSE)
  expect_equal(nrow(x), 5L)
  expect_equal(ncol(x), 4L)

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = TRUE, colNames = TRUE)
  expect_equal(nrow(x), 5L - 1L)
  expect_equal(ncol(x), 4L)


  ##############################################################
  ## FALSE FALSE FALSE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = FALSE, colNames = FALSE)
  expect_equal(nrow(x), 6L)
  expect_equal(ncol(x), 8L)

  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))

  ## NA rows
  expect_true(all(is.na(x[3, ])))



  ##############################################################
  ## FALSE FALSE TRUE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = FALSE, colNames = TRUE)
  expect_equal(nrow(x), 6L - 1L)
  expect_equal(ncol(x), 8L)


  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))

  ## NA rows
  expect_true(all(is.na(x[2, ])))



  ##############################################################
  ## FALSE TRUE FALSE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = TRUE, colNames = FALSE)
  expect_equal(nrow(x), 5L)
  expect_equal(ncol(x), 8L)

  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))




  ##############################################################
  ## FALSE TRUE TRUE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = TRUE, colNames = TRUE)
  expect_equal(nrow(x), 5L - 1L)
  expect_equal(ncol(x), 8L)


  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))



  ##############################################################
  ## TRUE FALSE FALSE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = FALSE, colNames = FALSE)
  expect_equal(nrow(x), 6L)
  expect_equal(ncol(x), 4L)

  ## NA rows
  expect_true(all(is.na(x[3, ])))



  ##############################################################
  ## TRUE FALSE TRUE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = FALSE, colNames = TRUE)
  expect_equal(nrow(x), 6L - 1L)
  expect_equal(ncol(x), 4L)


  ## NA rows
  expect_true(all(is.na(x[2, ])))
})






test_that("Version 4 fixes from Workbook Objects", {
  fl <- loadWorkbook(system.file("readTest.xlsx", package = "openxlsx"))


  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = TRUE, colNames = FALSE)
  expect_equal(nrow(x), 5L)
  expect_equal(ncol(x), 4L)

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = TRUE, colNames = TRUE)
  expect_equal(nrow(x), 5L - 1L)
  expect_equal(ncol(x), 4L)


  ##############################################################
  ## FALSE FALSE FALSE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = FALSE, colNames = FALSE)
  expect_equal(nrow(x), 6L)
  expect_equal(ncol(x), 8L)

  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))

  ## NA rows
  expect_true(all(is.na(x[3, ])))



  ##############################################################
  ## FALSE FALSE TRUE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = FALSE, colNames = TRUE)
  expect_equal(nrow(x), 6L - 1L)
  expect_equal(ncol(x), 8L)


  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))

  ## NA rows
  expect_true(all(is.na(x[2, ])))



  ##############################################################
  ## FALSE TRUE FALSE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = TRUE, colNames = FALSE)
  expect_equal(nrow(x), 5L)
  expect_equal(ncol(x), 8L)

  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))




  ##############################################################
  ## FALSE TRUE TRUE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = FALSE, skipEmptyRows = TRUE, colNames = TRUE)
  expect_equal(nrow(x), 5L - 1L)
  expect_equal(ncol(x), 8L)


  ## NA columns
  expect_true(all(is.na(x$X1)))
  expect_true(all(is.na(x$X2)))
  expect_true(all(is.na(x$X3)))
  expect_true(all(is.na(x$X7)))



  ##############################################################
  ## TRUE FALSE FALSE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = FALSE, colNames = FALSE)
  expect_equal(nrow(x), 6L)
  expect_equal(ncol(x), 4L)

  ## NA rows
  expect_true(all(is.na(x[3, ])))



  ##############################################################
  ## TRUE FALSE TRUE

  x <- read.xlsx(xlsxFile = fl, sheet = 4, skipEmptyCols = TRUE, skipEmptyRows = FALSE, colNames = TRUE)
  expect_equal(nrow(x), 6L - 1L)
  expect_equal(ncol(x), 4L)


  ## NA rows
  expect_true(all(is.na(x[2, ])))
})
