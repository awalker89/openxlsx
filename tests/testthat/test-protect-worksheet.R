

context("Protection")


test_that("Protection", {
  
  wb <- createWorkbook()
  addWorksheet(wb, "s1")
  addWorksheet(wb, "s2")
  
  
  protectWorksheet(wb, sheet = "s1", protect = TRUE, password = "abcdefghij", lockSelectingLockedCells = FALSE, lockSelectingUnlockedCells = FALSE, lockFormattingCells = TRUE, lockFormattingColumns = TRUE, lockPivotTables = TRUE)

  expect_true(wb$worksheets[[1]]$sheetProtection == "<sheetProtection password=\"FEF1\" selectLockedCells=\"0\" selectUnlockedCells=\"0\" formatCells=\"1\" formatColumns=\"1\" pivotTables=\"1\" sheet=\"1\"/>")
  
  protectWorksheet(wb, sheet = "s2", protect = TRUE)
  expect_true(wb$worksheets[[2]]$sheetProtection == "<sheetProtection sheet=\"1\"/>")
  protectWorksheet(wb, sheet = "s2", protect = FALSE)
  expect_true(wb$worksheets[[2]]$sheetProtection == "")
  
  
})
  

