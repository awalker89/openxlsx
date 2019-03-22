

context("Protection")


test_that("Protect Workbook", {
  wb <- createWorkbook()
  addWorksheet(wb, "s1")

  wb$protectWorkbook(password = "abcdefghij")

  expect_true(wb$workbookProtection == "<workbookProtection workbookPassword=\"FEF1\"/>")

  wb$protectWorkbook(protect = FALSE, password = "abcdefghij", lockStructure = TRUE, lockWindows = TRUE)
  expect_true(wb$workbookProtection == "")
})
