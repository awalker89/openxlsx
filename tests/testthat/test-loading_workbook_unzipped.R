



context("Load Unzipped Workbook Object")


test_that("Loading unzipped readTest.xlsx", {
  fl <- system.file("readTest.xlsx", package = "openxlsx")
  wb <- loadWorkbook(fl)

  ## make unzipped file & load
  tmp_dir <- file.path(tempdir(), paste(sample(LETTERS, 6), collapse = ""))
  if (dir.exists(tmp_dir)) unlink(tmp_dir, recursive = TRUE)
  dir.create(tmp_dir)

  unzip(zipfile = fl, exdir = tmp_dir)
  wb2 <- loadWorkbook(file = tmp_dir, isUnzipped = TRUE)

  expect_true(all.equal(wb, wb2))

  unlink(tmp_dir, recursive = TRUE)
})








""
