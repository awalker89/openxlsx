




context("Fill Merged Cells")



test_that("fill merged cells", {
  
  
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = "sheet1")
  writeData(wb = wb, sheet = 1, x = data.frame("A" = 1, "B" = 2))
  writeData(wb = wb, sheet = 1, x = 2, startRow = 2, startCol = 2)
  writeData(wb = wb, sheet = 1, x = 3, startRow = 2, startCol = 3)
  writeData(wb = wb, sheet = 1, x = 4, startRow = 2, startCol = 4)

  mergeCells(wb = wb, sheet = 1, cols = 2:4, rows = 1)
  
  tmp_file <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb = wb, file = tmp_file, overwrite = TRUE)
  
  expect_equal(names(read.xlsx(tmp_file, fillMergedCells = FALSE)), c("A", "B", "X3", "X4"))
  expect_equal(names(read.xlsx(tmp_file, fillMergedCells = TRUE)), c("A", "B", "B", "B"))
  
  
})




