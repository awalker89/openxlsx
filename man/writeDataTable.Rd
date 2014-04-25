\name{writeDataTable}
\alias{writeDataTable}
\title{Write to a worksheet and format as a table}
\usage{
writeDataTable(wb, sheet, x, startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleMedium2")
}
\arguments{
  \item{wb}{A Workbook object containing a worksheet.}

  \item{sheet}{The worksheet to write to. Can be the
  worksheet index or name.}

  \item{x}{A dataframe.}

  \item{startCol}{A vector specifiying the starting
  columns(s) to write df}

  \item{startRow}{A vector specifiying the starting row(s)
  to write df}

  \item{xy}{An alternative to specifying startCol and
  startRow individually. A vector of the form c(startCol,
  startRow)}

  \item{colNames}{If TRUE, column names of x are written.}

  \item{rowNames}{If TRUE, data.frame row names of x are
  written.}

  \item{tableStyle}{Any excel table style name.}
}
\description{
Write to a worksheet and format as a table
}
\details{
columns of x with class Date, POSIXct of POSIXt are
automatically styled as dates.
}
\examples{
## see package vignette for further examples.

wb <- createWorkbook("Edgar Anderson")

addWorksheet(wb, "S1")
addWorksheet(wb, "S2")
addWorksheet(wb, "S3")

## write data formatted as excel table with table filters
# default table style is "TableStyleMedium2"
writeDataTable(wb, "S1", x = iris)

writeDataTable(wb, "S2", x = mtcars, xy = c("B", 3), rowNames=TRUE, tableStyle="TableStyleLight9")

writeDataTable(wb, "S3", x = mtcars, startRow = 4, rowNames=TRUE, tableStyle="TableStyleMedium17")

saveWorkbook(wb, "writeDataTableExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{addWorksheet}}

\code{\link{writeData}}
}

