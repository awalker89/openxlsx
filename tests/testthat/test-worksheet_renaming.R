




context("Renaming worksheets.")



test_that("Can rename worksheets under all conditions", {
  tempFile <- file.path(tempdir(), "renaming.xlsx")
  wb <- createWorkbook()
  addWorksheet(wb, "sheet 1")
  addWorksheet(wb, "sheet 2")
  addWorksheet(wb, "sheet 3")
  addWorksheet(wb, "sheet 4")
  addWorksheet(wb, "sheet 5")

  renameWorksheet(wb, sheet = 2, "THis is SHEET 2")
  expect_equal(names(wb), c("sheet 1", "THis is SHEET 2", "sheet 3", "sheet 4", "sheet 5"))


  renameWorksheet(wb, sheet = "THis is SHEET 2", "THis is STILL SHEET 2")
  expect_equal(names(wb), c("sheet 1", "THis is STILL SHEET 2", "sheet 3", "sheet 4", "sheet 5"))


  renameWorksheet(wb, sheet = 5, "THis is SHEET 5")
  expect_equal(names(wb), c("sheet 1", "THis is STILL SHEET 2", "sheet 3", "sheet 4", "THis is SHEET 5"))

  renameWorksheet(wb, sheet = 5, "THis is STILL SHEET 5")
  expect_equal(names(wb), c("sheet 1", "THis is STILL SHEET 2", "sheet 3", "sheet 4", "THis is STILL SHEET 5"))


  renameWorksheet(wb, sheet = 2, "Sheet 2")
  expect_equal(names(wb), c("sheet 1", "Sheet 2", "sheet 3", "sheet 4", "THis is STILL SHEET 5"))

  renameWorksheet(wb, sheet = 5, "Sheet 5")
  expect_equal(names(wb), c("sheet 1", "Sheet 2", "sheet 3", "sheet 4", "Sheet 5"))


  ## re-ordering
  worksheetOrder(wb) <- c(4, 3, 2, 5, 1)
  saveWorkbook(wb, tempFile, overwrite = TRUE)

  wb <- loadWorkbook(file = tempFile)
  renameWorksheet(wb, sheet = 2, "THIS is SHEET 3")

  wb <- loadWorkbook(tempFile)
  renameWorksheet(wb, sheet = "Sheet 5", "THIS is NOW SHEET 5")

  expect_equal(names(wb), c("sheet 4", "sheet 3", "Sheet 2", "THIS is NOW SHEET 5", "sheet 1"))

  names(wb)[[1]] <- "THIS IS NOW SHEET 4"
  expect_equal(names(wb), c("THIS IS NOW SHEET 4", "sheet 3", "Sheet 2", "THIS is NOW SHEET 5", "sheet 1"))


  unlink(tempFile, recursive = TRUE, force = TRUE)
})
