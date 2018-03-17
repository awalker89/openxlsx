

context("Date/Time Conversions")



test_that("convert to date", {
  
  dates <- as.Date("2015-02-07")  + -10:10
  origin <- 25569L
  n <- as.integer(dates) + origin
  
  expect_equal(convertToDate(n), dates)
  
})



test_that("convert to datetime", {
  
  
  x <- 43037 + 2 / 1440
  expect_equal(object = convertToDateTime(x,tx=Sys.timezone()), expected = as.POSIXct("2017-10-29 00:02:00",tz=Sys.timezone()))
  
  x <- 43037 + 2/1440 + 1/86400
  expect_equal(object = convertToDateTime(x,tx=Sys.timezone()), expected = as.POSIXct("2017-10-29 00:02:01",tz=Sys.timezone()))
  
  x <- 43037 + 2.50 / 1440
  expect_equal(object = convertToDateTime(x,tx=Sys.timezone()), expected = as.POSIXct("2017-10-29 00:02:30",tz=Sys.timezone()))
  
  x <- 43037 + 2/1440 + 12.12/86400
  x_datetime<-convertToDateTime(x,tx="UTC")
  attr(x_datetime, "tzone") <- "UTC"
  
  x_test<-structure(1509195732.12+9*3600, class = c("POSIXct", "POSIXt"))
  attr(x_test, "tzone") <- "UTC"
  
  expect_equal(object = x_datetime , expected = x_test)
  
  
  
})