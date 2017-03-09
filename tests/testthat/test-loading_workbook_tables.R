

context("Load Workbook Object")


test_that("Tables loaded correctly", {
  
  wb <- loadWorkbook(system.file("loadExample.xlsx", package = "openxlsx"))

  expect_equal(unname(attr(wb$tables, "tableName")), c("Table2", "Table3"))
  expect_equal(names(attr(wb$tables, "tableName")), c("A1:E51", "A1:K30"))
  expect_equal(attr(wb$tables, "sheet"), c(1, 3))
  
  expect_equal(wb$worksheets[[1]]$tableParts, "<tablePart r:id=\"rId3\"/>", check.attributes = FALSE)
  expect_equal(unname(attr(wb$worksheets[[1]]$tableParts, "tableName")), "Table2")
  expect_equal(names(attr(wb$worksheets[[1]]$tableParts, "tableName")), "A1:E51")
  
  expect_equal(wb$worksheets[[3]]$tableParts, "<tablePart r:id=\"rId4\"/>", check.attributes = FALSE)
  expect_equal(unname(attr(wb$worksheets[[3]]$tableParts, "tableName")), "Table3")
  expect_equal(names(attr(wb$worksheets[[3]]$tableParts, "tableName")), "A1:K30")
  
  
  ## now remove a table
  expect_equal(unname(getTables(wb, 1)), "Table2", check.attributes = FALSE)
  expect_equal(unname(getTables(wb, 3)), "Table3", check.attributes = FALSE)
  
  
  
  expect_equal(getTables(wb, sheet = 1), character(0), check.attributes = FALSE)
  expect_equal(length(wb$worksheets[[1]]$tableParts), 0)
  expect_equal(wb$worksheets[[1]]$tableParts, character(0), check.attributes = FALSE)
  
  expect_equal(wb$worksheets[[3]]$tableParts, "<tablePart r:id=\"rId4\"/>", check.attributes = FALSE)
  expect_equal(unname(attr(wb$worksheets[[3]]$tableParts, "tableName")), "Table3")
  expect_equal(names(attr(wb$worksheets[[3]]$tableParts, "tableName")), "A1:K30")
  
  
  expect_error(removeTable(wb, sheet = 1, table = "Table2"), regexp = "table 'Table2' does not exist")

  
})