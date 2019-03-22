
context("Workbook properties")

test_that("Workbook properties", {

  ## check creator
  wb <- createWorkbook(creator = "Alex", title = "title here", subject = "this & that", category = "some category")

  expect_true("<dc:creator>Alex</dc:creator>" %in% wb$core)
  expect_true("<dc:title>title here</dc:title>" %in% wb$core)
  expect_true("<dc:subject>this &amp; that</dc:subject>" %in% wb$core)
  expect_true("<cp:category>some category</cp:category>" %in% wb$core)

  fl <- tempfile(fileext = ".xlsx")
  wb <- write.xlsx(
    x = iris, file = fl,
    creator =
      "Alex 2", title =
      "title here 2", subject =
      "this & that 2", category = "some category 2"
  )

  expect_true("<dc:creator>Alex 2</dc:creator>" %in% wb$core)
  expect_true("<dc:title>title here 2</dc:title>" %in% wb$core)
  expect_true("<dc:subject>this &amp; that 2</dc:subject>" %in% wb$core)
  expect_true("<cp:category>some category 2</cp:category>" %in% wb$core)

  ## maintain on load
  wb_loaded <- loadWorkbook(fl)
  expect_equal(object = wb_loaded$core, expected = paste0(wb$core, collapse = ""))
})
