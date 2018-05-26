






context("Writing Posixct")



test_that("Writing Posixct wikth writeData & writeDataTable", {
  
  options("openxlsx.datetimeFormat" = "dd/mm/yy hh:mm")
  
  tstart <- strptime("30/05/2017 08:30", "%d/%m/%Y %H:%M", tz="CET")
  TimeDT <- c(0,5,10,15,30,60,120,180,240,480,720,1440)*60 + tstart
  df <- data.frame(TimeDT, TimeTxt = format(TimeDT,"%Y-%m-%d %H:%M"))
  
  wb <- createWorkbook()
  addWorksheet(wb, "writeData")
  addWorksheet(wb, "writeDataTable")
  
  writeData(wb, "writeData", df, startCol = 2, startRow = 3, rowNames = FALSE)
  writeDataTable(wb, "writeDataTable", df, startCol = 2, startRow = 3)
  
  wd <- wb$worksheets[[1]]$sheet_data$v
  wdt <- wb$worksheets[[2]]$sheet_data$v
  
  
  expected <- c("0", "1", "42885.354166666664", "2", "42885.357638888891", 
                "3", "42885.361111111117", "4", "42885.364583333336", "5", "42885.375000000007", 
                "6", "42885.395833333336", "7", "42885.437500000007", "8", "42885.479166666664", 
                "9", "42885.520833333336", "10", "42885.687500000007", "11", 
                "42885.854166666664", "12", "42886.354166666664", "13")
  
  expect_equal(object = wd, expected = expected)
  expect_equal(object = wdt, expected = expected)
  expect_equal(object = wd, expected = wdt)
  
  options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
  
})

