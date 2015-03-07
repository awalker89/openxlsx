


context("Date Conversion Testing")



test_that("convert to datetime", {
 
  dates <- as.Date("2015-02-07")  + -10:10
  origin <- 25569L
  n <- as.integer(dates) + origin
  
  expect_equal(convertToDate(n), dates)
  
})




