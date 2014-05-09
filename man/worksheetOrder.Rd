\name{worksheetOrder}
\alias{worksheetOrder}
\alias{worksheetOrder<-}
\title{Order of worksheets in xlsx file}
\usage{
worksheetOrder(wb)

worksheetOrder(wb) <- value
}
\arguments{
  \item{wb}{A workbook object}

  \item{value}{Vector specifying order to write worksheets
  to file}
}
\description{
Order of worksheets in xlsx file
}
\details{
This function does not reorder the worksheets within the
workbook object, it simply shuffles the order when writing
to file.
}
\examples{
wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = "Sheet 1", gridLines=F)
writeDataTable(wb = wb, sheet = 1, x = iris)

addWorksheet(wb = wb, sheetName = "mtcars (Sheet 2)", gridLines=F)
writeData(wb = wb, sheet = 2, x = mtcars)

addWorksheet(wb = wb, sheetName = "Sheet 3", gridLines=F)
writeData(wb = wb, sheet = 3, x = Formaldehyde)

worksheetOrder(wb)
sheets(wb)
worksheetOrder(wb) <- c(1,3,2) # switch position of sheets 2 & 3
writeData(wb, 2, 'This is still the "mtcars" worksheet', startCol = 15)
worksheetOrder(wb)
sheets(wb)

saveWorkbook(wb, "worksheetOrderExample.xlsx",  overwrite = TRUE)
}

