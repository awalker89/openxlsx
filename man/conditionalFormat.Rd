\name{conditionalFormat}
\alias{conditionalFormat}
\title{Add conditional formatting to cells}
\usage{
conditionalFormat(wb, sheet, cols, rows, rule, style = NULL)
}
\arguments{
  \item{wb}{A workbook object}

  \item{sheet}{A name or index of a worksheet}

  \item{cols}{Columns to apply conditional formattiong}

  \item{rows}{Rows to apply conditional formattiong}

  \item{rule}{The condition under which to apply the
  formatting.  See details.}

  \item{style}{A Style to apply to those cells that
  satisify the rule. A Style object returned from
  createStyle()}
}
\description{
Add conditional formatting to cells
}
\details{
Conditional format will be applied to all cells within
region specified by rows X cols. valid operators are "<",
"<=", ">", ">=", "==", "!=". Default style given by:
createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
}
\examples{
wb <- createWorkbook()

## Add a worksheet and data
addWorksheet(wb, "Worksheet Name")
writeData(wb, 1, matrix(runif(100), ncol = 10), xy = c(1,1), colNames=FALSE)
writeData(wb, 1, matrix(sample(LETTERS, 100, rep = TRUE), ncol = 10), xy = c(1, 11), colNames=FALSE)

## Create a style
conditionalStyle <- createStyle(bgFill = "#FFFF00")

## Conditionally format cells:Rows 10 to 20 columns A to J (1:10)
conditionalFormat(wb, 1, cols = 1:10, rows = 1:10, rule = "< 0.5", style = conditionalStyle)
conditionalFormat(wb, 1, cols = 1:10, rows = 11:20, rule = '== "A"')  # default style

## Save workbook
saveWorkbook(wb, "conditionalFormatExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{createStyle}}
}

