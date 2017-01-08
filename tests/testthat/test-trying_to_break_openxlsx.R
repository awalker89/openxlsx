


context("Images and Tables.")


test_that("Images and Tables - reordering and removing", {

  options('stringsAsFactors' = FALSE)
  tempFile <- file.path(tempdir(), "break.xlsx")

  getPlot <- function(i){
    n <- 5000
    plot(1:n, rnorm(n))
    title(main = sprintf("Plot for Sheet: %s", i))
  }

  df1 <- iris[1:5, 1:4]
  df2 <- mtcars
  
  
  df3 <- data.frame("Date" = Sys.Date()-0:10,
                   "Logical" = sample(c(TRUE, FALSE), 1, replace = TRUE),
                   "Currency" = as.numeric(-5:5)*100,
                   "Accounting" = as.numeric(-5:5),
                   "hLink" = "http://cran.r-project.org/", 
                   "Percentage" = seq(-5, 5, length.out=11),
                   "TinyNumber" = runif(11) / 1E9, stringsAsFactors = FALSE)
    
  df3U <- df3
    
  class(df3$Currency) <- "currency"
  class(df3$Accounting) <- "accounting"
  class(df3$hLink) <- "hyperlink"
  class(df3$Percentage) <- "percentage"
  class(df3$TinyNumber) <- "scientific"
  

  df4 <- data.frame("X" = 1:10000, "Y" = sample(LETTERS, size = 10000, replace = TRUE))
  df5 <- USJudgeRatings
    
  hs <- createStyle(fontColour = "blue", textRotation = 45)
  
  
  wb <- createWorkbook()
  expect_equal(names(wb), character(0))
  
  addWorksheet(wb = wb, sheetName = "Sheet 1", gridLines = FALSE, tabColour = "red", zoom = 75)
  writeDataTable(wb, sheet = 1, x = df1, startCol = 7, startRow = 10, tableName = "Sheet1Table1")
  expect_equal(names(wb),  "Sheet 1")
  
  
  addWorksheet(wb, sheetName = "Sheet 2", tabColour = "purple")
  writeDataTable(wb, sheet = "Sheet 2", x = df2, startCol = 2, startRow = 2, rowNames = TRUE)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2"))
  
  

  addWorksheet(wb, sheetName = "Sheet 3", tabColour = "green")
  writeDataTable(wb, sheet = 3, x = df3, startCol = 1, startRow = 1)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3"))
  
  addWorksheet(wb, sheetName = "Sheet 4", tabColour = "orange")
  writeDataTable(wb, sheet = 4, x = df4)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
  
  addWorksheet(wb, sheetName = "Sheet 5", tabColour = "yellow")
  writeData(wb, sheet = "Sheet 5", x = df5, rowNames = TRUE)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5"))
  

  
  worksheetOrder(wb) <- c(1, 3, 5, 4, 2)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4", "Sheet 5"))
  
  ## save and load 1
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)

  wb <- loadWorkbook(tempFile)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 3", "Sheet 5", "Sheet 4", "Sheet 2"))


  expect_equal(df1,  read.xlsx(wb, sheet = 1))
  expect_equal(df1,  read.xlsx(wb, sheet = "Sheet 1"))
  expect_equal(df1,  read.xlsx(tempFile, sheet = 1))
  expect_equal(df1,  read.xlsx(tempFile, sheet = "Sheet 1"))
  
  
  expect_equal(df3U,  read.xlsx(wb, sheet = 2, detectDates = TRUE))
  expect_equal(df3U,  read.xlsx(wb, sheet = "Sheet 3", detectDates = TRUE))
  expect_equal(df3U,  read.xlsx(tempFile, sheet = 2, detectDates = TRUE))
  expect_equal(df3U,  read.xlsx(tempFile, sheet = "Sheet 3", detectDates = TRUE))
    
  
  expect_equal(df5,  read.xlsx(wb, sheet = 3, rowNames = TRUE))
  expect_equal(df5,  read.xlsx(wb, sheet = "Sheet 5", rowNames = TRUE))
  expect_equal(df5,  read.xlsx(tempFile, sheet = 3, rowNames = TRUE))
  expect_equal(df5,  read.xlsx(tempFile, sheet = "Sheet 5", rowNames = TRUE))
  
  
  expect_equal(df4,  read.xlsx(wb, sheet = 4))
  expect_equal(df4,  read.xlsx(wb, sheet = "Sheet 4"))
  expect_equal(df4,  read.xlsx(tempFile, sheet = 4))
  expect_equal(df4,  read.xlsx(tempFile, sheet = "Sheet 4"))
  

  expect_equal(df2,  read.xlsx(wb, sheet = 5, rowNames = TRUE))
  expect_equal(df2,  read.xlsx(wb, sheet = "Sheet 2", rowNames = TRUE))
  expect_equal(df2,  read.xlsx(tempFile, sheet = 5, rowNames = TRUE))
  expect_equal(df2,  read.xlsx(tempFile, sheet = "Sheet 2", rowNames = TRUE))
  


  ## remove "Sheet 5" by index (3)
  removeWorksheet(wb, sheet = 3)
  expect_equal(names(wb),  c("Sheet 1", "Sheet 3", "Sheet 4", "Sheet 2"))

  ## remove sheet "Sheet 4"
  removeWorksheet(wb, sheet = "Sheet 4")
  expect_equal(names(wb),  c("Sheet 1", "Sheet 3", "Sheet 2"))


  # ## Introduce some images
  # getPlot(1)
  # insertPlot(wb = wb, sheet = "Sheet 1", startCol = 14, startRow = 3)
  # 
  # getPlot(2)
  # insertPlot(wb = wb, sheet = "Sheet 2", startCol = 14, startRow = 3)
  # 
  # getPlot(3)
  # insertPlot(wb = wb, sheet = "Sheet 3", startCol = 14, startRow = 3)
  # 
  # 
  # expect_true(any(grepl("image1", wb$drawings_rels[[1]])))
  # expect_true(any(grepl("image3", wb$drawings_rels[[2]])))
  # expect_true(any(grepl("image2", wb$drawings_rels[[3]])))



  ## put back to original order
  worksheetOrder(wb) <- c(1, 3, 2)
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)

  wb <- loadWorkbook(file = tempFile)


  # ## drawings added in order
  # expect_true(any(grepl("image1", wb$drawings_rels[[1]])))
  # expect_true(any(grepl("image2", wb$drawings_rels[[2]])))
  # expect_true(any(grepl("image3", wb$drawings_rels[[3]])))


  ## Introduce some more images
  getPlot("1_2")
  insertPlot(wb = wb, sheet = "Sheet 1", startCol = 14, startRow = 25)
  
  getPlot("2_2")
  insertPlot(wb = wb, sheet = "Sheet 2", startCol = 14, startRow = 25)
  
  
  getPlot("3_2")
  insertPlot(wb = wb, sheet = "Sheet 3", startCol = 14, startRow = 25)
  
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)
  wb <- loadWorkbook(tempFile)

  worksheetOrder(wb) <- c(3, 2, 1)
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)
  wb <- loadWorkbook(tempFile)


  hl <- rep("http://google.com.au", 5)
  names(hl) <- sprintf("Link to google %s", 1:5)
  class(hl) <- "hyperlink"  
  writeData(wb, "Sheet 1", hl)
    
  ## Add in some column widths

  setColWidths(wb, sheet = 1, cols = 1:50, widths = "auto")
  worksheetOrder(wb) <- c(3, 2, 1)
  removeWorksheet(wb, sheet = "Sheet 2")
  
  saveWorkbook(wb, file = tempFile, overwrite = TRUE)
  wb <- loadWorkbook(tempFile)
    
  expect_equal(names(wb),  c("Sheet 1", "Sheet 3"))
  expect_equal(df1, read.xlsx(tempFile, sheet = 1, startRow = 10))
  expect_equal(df3U,  read.xlsx(tempFile, sheet = 2, detectDates = TRUE))

  expect_equal(df1, read.xlsx(wb, sheet = 1, startRow = 10))
  expect_equal(df3U,  read.xlsx(wb, sheet = 2, detectDates = TRUE))
  
    
  unlink(tempFile, recursive = TRUE, force = TRUE)
  rm(wb)
  
  
})


