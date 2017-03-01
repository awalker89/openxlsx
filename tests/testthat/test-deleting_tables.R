

context(desc = "Deleting tables from worksheets")

test_that("Deleting a Table Object", {
  
  

  wb <- createWorkbook()
  addWorksheet(wb, sheetName = "Sheet 1")
  addWorksheet(wb, sheetName = "Sheet 2")
  writeDataTable(wb, sheet = "Sheet 1", x = iris, tableName = "iris")
  writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
  
  
  ###################################################################################
  ## Get table
  
  expect_equal(length(getTables(wb, sheet = 1)), 2L)
  expect_equal(length(getTables(wb, sheet = "Sheet 1")), 2L)
  
  expect_equal(length(getTables(wb, sheet = 2)), 0)
  expect_equal(length(getTables(wb, sheet = "Sheet 2")), 0)
  
  expect_error(getTables(wb, sheet = 3))
  expect_error(getTables(wb, sheet = "Sheet 3"))  
  
  expect_equal(getTables(wb, sheet = 1), c("iris", "mtcars"), check.attributes = FALSE)
  expect_equal(getTables(wb, sheet = "Sheet 1"), c("iris", "mtcars"), check.attributes = FALSE)
  
  expect_equal(attr(getTables(wb, sheet = 1), "refs"), c("A1:E151", "J1:T33"))
  expect_equal(attr(getTables(wb, sheet = "Sheet 1"), "refs"), c("A1:E151", "J1:T33"))
  
  expect_equal(length(wb$tables), 2L)
  
  
  
  
  
  ###################################################################################
  ## Deleting a worksheet
  
  removeWorksheet(wb, 1)
  expect_equal(length(wb$tables), 2L)
  expect_equal(length(getTables(wb, sheet = 1)), 0)
  
  expect_equal(attr(wb$tables, "tableName"), c("iris_openxlsx_deleted", "mtcars_openxlsx_deleted"))
  expect_equal(attr(wb$tables, "sheet"), c(1, 1))
  
  
  
  
  ###################################################################################
  ## write same tables again
  
  writeDataTable(wb, sheet = 1, x = iris, tableName = "iris")
  writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
  
  
  expect_equal(length(getTables(wb, sheet = 1)), 2L)
  expect_equal(length(getTables(wb, sheet = "Sheet 2")), 2L)
  
  expect_error(getTables(wb, sheet = 2))
  expect_error(getTables(wb, sheet = "Sheet 1"))  
  
  expect_equal(getTables(wb, sheet = 1), c("iris", "mtcars"), check.attributes = FALSE)
  expect_equal(getTables(wb, sheet = "Sheet 2"), c("iris", "mtcars"), check.attributes = FALSE)
  
  expect_equal(attr(getTables(wb, sheet = 1), "refs"), c("A1:E151", "J1:T33"))
  expect_equal(attr(getTables(wb, sheet = "Sheet 2"), "refs"), c("A1:E151", "J1:T33"))
  
  expect_equal(length(wb$tables), 4L)
  
  
  ###################################################################################
  ## removeTable
  
  ## remove iris and re-write it
  removeTable(wb = wb, sheet = 1, table = "iris")
  
  expect_equal(length(wb$tables), 4L)
  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId6\"/>", check.attributes = FALSE)
  expect_equal(attr(wb$worksheets[[1]]$tableParts, "tableName"), "mtcars")
  
  expect_equal(attr(wb$tables, "tableName"), c("iris_openxlsx_deleted",
                                               "mtcars_openxlsx_deleted",
                                               "iris_openxlsx_deleted",
                                               "mtcars"))
  
  
  
  
  ## removeTable clears table object and all data
  writeDataTable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
  expect_equal(wb$worksheets[[1]]$tableParts, c("<tablePart r:id=\"rId6\"/>", "<tablePart r:id=\"rId7\"/>"), check.attributes = FALSE)
  expect_equal(attr(wb$worksheets[[1]]$tableParts, "tableName"), c("mtcars", "iris"))
  
  
  removeTable(wb = wb, sheet = 1, table = "iris")
  
  expect_equal(length(wb$tables), 5L)
  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId6\"/>", check.attributes = FALSE)
  expect_equal(attr(wb$worksheets[[1]]$tableParts, "tableName"), "mtcars")
  
  expect_equal(attr(wb$tables, "tableName"), c("iris_openxlsx_deleted",
                                               "mtcars_openxlsx_deleted",
                                               "iris_openxlsx_deleted",
                                               "mtcars",
                                               "iris_openxlsx_deleted"))
  
  
  expect_equal(getTables(wb, sheet = 1), "mtcars", check.attributes = FALSE)
  
  
  
})





