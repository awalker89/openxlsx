

context("write.xlsx vector arguments")

test_that("Writing then reading returns identical data.frame 1", {
  tmp_file <- file.path(tempdir(), "xlsx_vector_args.xlsx")

  df1 <- data.frame(1:2)
  df2 <- data.frame(1:3)
  x <- list(df1, df2)

  write.xlsx(
    file = tmp_file,
    x = x,
    gridLines = c(F, T),
    sheetName = c("a", "b"),
    zoom = c(50, 90),
    tabColour = c("red", "blue")
  )

  wb <- loadWorkbook(tmp_file)

  expect_equal(object = getSheetNames(tmp_file), expected = c("a", "b"))
  expect_equal(object = names(wb), expected = c("a", "b"))

  expect_true(object = grepl('rgb="FFFF0000"', wb$worksheets[[1]]$sheetPr))
  expect_true(object = grepl('rgb="FF0000FF"', wb$worksheets[[2]]$sheetPr))

  expect_true(object = grepl('zoomScale="50"', wb$worksheets[[1]]$sheetViews))
  expect_true(object = grepl('zoomScale="90"', wb$worksheets[[2]]$sheetViews))

  expect_true(object = grepl('showGridLines="0"', wb$worksheets[[1]]$sheetViews))
  expect_true(object = grepl('showGridLines="1"', wb$worksheets[[2]]$sheetViews))

  expect_equal(read.xlsx(tmp_file, sheet = 1), df1)
  expect_equal(read.xlsx(tmp_file, sheet = 2), df2)

  unlink(tmp_file, recursive = TRUE, force = TRUE)
})
