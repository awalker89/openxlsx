






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
  
  
  expected <- c("0", "1", "42885.35416666666424135", "2", "42885.35763888889050577", 
                "3", "42885.36111111111677019", "4", "42885.36458333333575865", "5", "42885.37500000000727596", 
                "6", "42885.39583333333575865", "7", "42885.43750000000727596", "8", "42885.47916666666424135", 
                "9", "42885.52083333333575865", "10", "42885.68750000000727596", "11", 
                "42885.85416666666424135", "12", "42886.35416666666424135", "13")
  
  expect_equal(object = wd, expected = expected)
  expect_equal(object = wdt, expected = expected)
  expect_equal(object = wd, expected = wdt)
  
  options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
  
})

