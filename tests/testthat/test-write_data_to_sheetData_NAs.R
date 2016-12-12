



context("Writing NA to SheetData")



test_that("Writing to sheetData with keepNA = FALSE", {
  
  
  a <- head(iris)
  a[2,2] <- NA
  a[3,5] <- NA
  a[5,1] <- NA
  a[5,5] <- NA
  
  wb <- createWorkbook()
  addWorksheet(wb, "Sheet 1")
  writeData(wb, 1, a, keepNA = FALSE)
  
  sheet_data <- wb$worksheets[[1]]$sheetData
  
  sheet_v <- unlist(unname(lapply(sheet_data, "[[", "v")))
  sheet_t <- unlist(unname(lapply(sheet_data, "[[", "t")))
  sheet_f <- unlist(unname(lapply(sheet_data, "[[", "f")))
  sheet_r <- unlist(unname(lapply(sheet_data, "[[", "r")))
  
  sheet_row <- as.integer(gsub("[A-Z]", "", sheet_r))
  sheet_col <- convertFromExcelRef(sheet_r)
  sheet_v <- as.numeric(sheet_v)
  
  sheet_v <- sheet_v[!is.na(sheet_v)]
  sheet_t <- sheet_t[!is.na(sheet_t)]
  sheet_f <- sheet_f[!is.na(sheet_f)]
  sheet_r <- sheet_r[!is.na(sheet_r)]
  sheet_row <- sheet_row[!is.na(sheet_row)]
  sheet_col <- sheet_col[!is.na(sheet_col)]

  n_values <- prod(dim(a)) + ncol(a) ## number of NAs
  
  expect_length(sheet_row, n_values)
  expect_length(sheet_col, n_values)
  expect_length(sheet_t, n_values - 4)
  expect_length(sheet_v, n_values - 4)
  expect_length(sheet_f, 0)

  ## check sheetData vector
  expect_equal(sheet_row, rep(1:7, each = 5))
  expect_equal(sheet_col, rep(1:5, times = 7))
  
  ## header types
  expect_equal(sheet_t[1:5], rep("s", ncol(a)))
  
  ## data.frame t & v
  expect_equal(sheet_t[6:n_values], c("n", "n", "n", "n", "s", "n", "n", "n", "s", "n", "n", "n", 
                                      "n", "n", "n", "n", "n", "s", "n", "n", "n", "n", "n", "n", "n", 
                                      "s", NA, NA, NA, NA))
  
  expect_equal(sheet_v[1:5], 0:4)

  expected_v <- c(5.1, 3.5, 1.4, 0.2, 5, 4.9, 1.4, 0.2, 5, 4.7, 3.2, 1.3, 0.2, 
                  4.6, 3.1, 1.5, 0.2, 5, 3.6, 1.4, 0.2, 5.4, 3.9, 1.7, 0.4, 5, 
                  NA, NA, NA, NA)
  
  expect_equal(sheet_v[6:n_values], expected_v)
  
  
  
  
})





