


test_that("Validate Table Names", {
  
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")

  ## case
  expect_equal(wb$validate_table_name("test"), "test")
  expect_equal(wb$validate_table_name("TEST"), "test")
  expect_equal(wb$validate_table_name("Test"), "test")
  
  ## length
  expect_error(wb$validate_table_name(paste(sample(LETTERS, size = 300, replace = TRUE), collapse = "")), regexp = 'tableName must be less than 255 characters')
  
  ## look like cell ref
  expect_error(wb$validate_table_name("R1C2"), regexp = 'tableName cannot be the same as a cell reference, such as R1C1', fixed = TRUE)
  expect_error(wb$validate_table_name("A1"), regexp = 'tableName cannot be the same as a cell reference', fixed = TRUE)
  
  expect_error(wb$validate_table_name("R06821C9682"), regexp = 'tableName cannot be the same as a cell reference, such as R1C1', fixed = TRUE)
  expect_error(wb$validate_table_name("ABD918751"), regexp = 'tableName cannot be the same as a cell reference', fixed = TRUE)
  
  expect_error(wb$validate_table_name("A$100"), regexp = "'$' character cannot exist in a tableName", fixed = TRUE)
  expect_error(wb$validate_table_name("A12$100"), regexp = "'$' character cannot exist in a tableName", fixed = TRUE)
  
  tbl_nm <- "性別"
  expect_equal(wb$validate_table_name(tbl_nm), tbl_nm)
  
  
})





test_that("Existing Table Names", {
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  ## Existing names - case in-sensitive
  
  writeDataTable(wb, sheet = 1, x = head(iris), tableName = "Table1")
  expect_error(wb$validate_table_name("Table1"), regexp = "Table with name 'table1' already exists", fixed = TRUE)
  expect_error(writeDataTable(wb, sheet = 1, x = head(iris), tableName = "Table1", startCol = 10), regexp = "Table with name 'table1' already exists", fixed = TRUE)
  
  expect_error(wb$validate_table_name("TABLE1"), regexp = "Table with name 'table1' already exists", fixed = TRUE)
  expect_error(writeDataTable(wb, sheet = 1, x = head(iris), tableName = "TABLE1", startCol = 20), regexp = "Table with name 'table1' already exists", fixed = TRUE)
  
  expect_error(wb$validate_table_name("table1"), regexp = "Table with name 'table1' already exists", fixed = TRUE)
  expect_error(writeDataTable(wb, sheet = 1, x = head(iris), tableName = "table1", startCol = 30), regexp = "Table with name 'table1' already exists", fixed = TRUE)
  
  

  
})


