




context("Writing over tables")



test_that("writeDataTable over tables", {
  overwrite_table_error <- "Cannot overwrite existing table with another table"
  df1 <- data.frame("X" = 1:10)

  wb <- createWorkbook()
  addWorksheet(wb, "Sheet1")

  ## table covers rows 4->10 and cols 4->8
  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 4)

  ## should all run without error
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 3, startRow = 2)
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 9, startRow = 2)
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 4, startRow = 11)
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 5, startRow = 11)
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 6, startRow = 11)
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 7, startRow = 11)
  writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 8, startRow = 11)
  writeDataTable(wb = wb, sheet = 1, x = head(iris, 2), startCol = 4, startRow = 1)



  ## Now error
  expect_error(writeDataTable(wb = wb, sheet = 1, x = df1, startCol = "H", startRow = 21), regexp = overwrite_table_error)
  expect_error(writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 3, startRow = 12), regexp = overwrite_table_error)
  expect_error(writeDataTable(wb = wb, sheet = 1, x = df1, startCol = 9, startRow = 12), regexp = overwrite_table_error)
  expect_error(writeDataTable(wb = wb, sheet = 1, x = df1, startCol = "i", startRow = 12), regexp = overwrite_table_error)


  ## more errors
  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris)), regexp = overwrite_table_error)
  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 21), regexp = overwrite_table_error)

  ## should work
  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 22)
  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 40)


  ## more errors
  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris, 2), startCol = 4, startRow = 38), regexp = overwrite_table_error)
  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris, 2), startCol = 4, startRow = 38, colNames = FALSE), regexp = overwrite_table_error)

  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = "H", startRow = 40), regexp = overwrite_table_error)
  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = "I", startRow = 40)
  writeDataTable(wb = wb, sheet = 1, x = head(iris)[, 1:3], startCol = "A", startRow = 40)

  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris, 2), startCol = 4, startRow = 38, colNames = FALSE), regexp = overwrite_table_error)
  expect_error(writeDataTable(wb = wb, sheet = 1, x = head(iris, 2), startCol = 1, startRow = 46, colNames = FALSE), regexp = overwrite_table_error)
})





test_that("writeData over tables", {
  overwrite_table_error <- "Cannot overwrite table headers. Avoid writing over the header row"
  df1 <- data.frame("X" = 1:10)

  wb <- createWorkbook()
  addWorksheet(wb, "Sheet1")

  ## table covers rows 4->10 and cols 4->8
  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 4)

  ## Anywhere on row 5 is fine
  for (i in 1:10)
    writeData(wb = wb, sheet = 1, x = head(iris), startRow = 5, startCol = i)

  ## Anywhere on col i is fine
  for (i in 1:10)
    writeData(wb = wb, sheet = 1, x = head(iris), startRow = i, startCol = "i")



  ## Now errors on headers
  expect_error(writeData(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 4), regexp = overwrite_table_error)
  writeData(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 5)
  writeData(wb = wb, sheet = 1, x = head(iris)[1:3])
  writeData(wb = wb, sheet = 1, x = head(iris, 2), startCol = 4)
  writeData(wb = wb, sheet = 1, x = head(iris, 2), startCol = 4, colNames = FALSE)


  ## Example of how this should be used
  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 30)
  writeData(wb = wb, sheet = 1, x = head(iris), startCol = 4, startRow = 31, colNames = FALSE)

  writeDataTable(wb = wb, sheet = 1, x = head(iris), startCol = 10, startRow = 30)
  writeData(wb = wb, sheet = 1, x = tail(iris), startCol = 10, startRow = 31, colNames = FALSE)

  writeDataTable(wb = wb, sheet = 1, x = head(iris)[, 1:3], startCol = 1, startRow = 30)
  writeData(wb = wb, sheet = 1, x = tail(iris), startCol = 1, startRow = 31, colNames = FALSE)
})
