\name{addWorksheet}
\alias{addWorksheet}
\title{Add a worksheet to a workbook}
\usage{
addWorksheet(wb, sheetName, gridLines = TRUE)
}
\arguments{
  \item{wb}{A Workbook object to attach the new worksheet}

  \item{sheetName}{A name for the new worksheet}

  \item{gridLines}{A logical. If TRUE, the worksheet grid
  lines will be shows, else they will be hidden. or
  hidden.}
}
\value{
XML tree
}
\description{
Add a worksheet to a workbook
}
\examples{
## Create a new workbook
wb <- createWorkbook("Fred")

## Add 3 worksheets
addWorksheet(wb, "Worksheet Name")
addWorksheet(wb, "This is worksheet 2")
addWorksheet(wb, "The third worksheet")

## Save workbook
saveWorkbook(wb, "addWorksheetExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}

