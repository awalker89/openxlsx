
context("Freeze Panes")

test_that("Freeze Panes", {
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstActiveRow = 3, firstActiveCol = 3)

  expected <- "<pane ySplit=\"2\" xSplit=\"2\" topLeftCell=\"C3\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)




  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstActiveRow = 1, firstActiveCol = 3)

  expected <- "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)




  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstActiveRow = 2, firstActiveCol = 1)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)




  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstActiveRow = 2, firstActiveCol = 4)

  expected <- "<pane ySplit=\"1\" xSplit=\"3\" topLeftCell=\"D2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)




  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstCol = TRUE)

  expected <- "<pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)




  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstRow = TRUE)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)




  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  freezePane(wb, 1, firstRow = TRUE, firstCol = TRUE)

  expected <- "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)


  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  addWorksheet(wb, "Sheet 2")
  addWorksheet(wb, "Sheet 3")
  addWorksheet(wb, "Sheet 4")
  addWorksheet(wb, "Sheet 5")
  addWorksheet(wb, "Sheet 6")
  addWorksheet(wb, "Sheet 7")

  freezePane(wb, sheet = 1, firstActiveRow = 3, firstActiveCol = 3)
  freezePane(wb, sheet = 2, firstActiveRow = 1, firstActiveCol = 3)
  freezePane(wb, sheet = 3, firstActiveRow = 2, firstActiveCol = 1)
  freezePane(wb, sheet = 4, firstActiveRow = 2, firstActiveCol = 4)
  freezePane(wb, sheet = 5, firstCol = TRUE)
  freezePane(wb, sheet = 6, firstRow = TRUE)
  freezePane(wb, sheet = 7, firstRow = TRUE, firstCol = TRUE)


  expected <- "<pane ySplit=\"2\" xSplit=\"2\" topLeftCell=\"C3\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[1]]$freezePane, expected)

  expected <- "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/>"
  expect_equal(wb$worksheets[[2]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/>"
  expect_equal(wb$worksheets[[3]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" xSplit=\"3\" topLeftCell=\"D2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[4]]$freezePane, expected)

  expected <- "<pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[5]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
  expect_equal(wb$worksheets[[6]]$freezePane, expected)

  expected <- "<pane ySplit=\"1\" xSplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\"/>"
  expect_equal(wb$worksheets[[7]]$freezePane, expected)
})
