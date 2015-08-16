



context("Named Regions")



test_that("Maintaining Named Regions on Load", {
 

  ## create named regions
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  
  ## specify region
  writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
  createNamedRegion(wb = wb,
                    sheet = 1,
                    name = "iris",
                    rows = 1:(nrow(iris)+1),
                    cols = 1:(ncol(iris)+1))
  
  
  ## using writeData 'name' argument
  writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
  
  out_file <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, out_file, overwrite = TRUE)
  
  expect_equal(object = getNamedRegions(wb), expected = getNamedRegions(out_file))
  
  df1 <- read.xlsx(wb, namedRegion = "iris")
  df2 <- read.xlsx(out_file, namedRegion = "iris")
  expect_equal(object = df1, expected = df2)
  
  
  
  
})








