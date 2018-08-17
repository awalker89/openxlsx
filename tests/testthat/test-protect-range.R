

context("Protection")


test_that("Protect Range", {
  
  wb <- createWorkbook()
  addWorksheet(wb, "s1")
  
  protectRange(wb, sheet = "s1", range = "A1", password = "abcdefghij", name = "myrange")
  
  expect_true(wb$worksheets[[1]]$protectedRanges == "<protectedRanges><protectedRange password=\"FEF1\" sqref=\"A1\" name=\"myrange\"/></protectedRanges>")
  
  # Multiple protected ranges are allowed
  protectRange(wb, sheet = "s1", range = "A2", password = "abcdefghij", name = "myrange2")
  
  expect_true(wb$worksheets[[1]]$protectedRanges == "<protectedRanges><protectedRange password=\"FEF1\" sqref=\"A1\" name=\"myrange\"/><protectedRange password=\"FEF1\" sqref=\"A2\" name=\"myrange2\"/></protectedRanges>")
  
})


