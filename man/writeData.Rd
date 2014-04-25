\name{writeData}
\alias{writeData}
\title{Write to a worksheet}
\usage{
writeData(wb, sheet, x, startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, headerStyle = NULL, borders = NULL,
  borderColour = "#4F81BD")
}
\arguments{
  \item{wb}{A Workbook object containing a worksheet.}

  \item{sheet}{The worksheet to write to. Can be the
  worksheet index or name.}

  \item{x}{Data to be written.}

  \item{startCol}{A vector specifiying the starting
  columns(s) to write df}

  \item{startRow}{A vector specifiying the starting row(s)
  to write df}

  \item{xy}{An alternative to specifying startCol and
  startRow individually.  A vector of the form c(startCol,
  startRow)}

  \item{colNames}{If TRUE, column names of x are written.}

  \item{rowNames}{If TRUE, data.frame row names of x are
  written.}

  \item{headerStyle}{Custom style to apply to column
  names.}

  \item{borders}{Either "surrounding", "columns" or "rows"
  or NULL.  If "surrounding", a border is drawn around the
  data.  If "rows", a surrounding border is drawn a border
  around each row. If "columns", a surrounding border is
  drawn with a border between each column.}

  \item{borderColour}{Colour of cell border.  A valid hex
  colour beginning with "#"}
}
\description{
Write data to worksheet with optional styling.
}
\examples{
## see package vignette for further examples.

wb <- createWorkbook()

## Add worksheets
addWorksheet(wb, "Cars")

x <- mtcars[1:6,]
writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
writeData(wb, "Cars", x, rowNames = TRUE, startCol = "O", startRow = 3,
         borders="surrounding", borderColour = NULL) ## black border

writeData(wb, "Cars", x, rowNames = TRUE,
         startCol = 2, startRow = 12, borders="columns")

writeData(wb, "Cars", x, rowNames = TRUE,
         startCol="O", startRow = 12, borders="rows", borderColour = "#4F81BD")

## header styles
hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "Italic",
                   border = "Bottom", borderColour = "#4F81BD")

writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
          startCol="B", startRow = 23, borders="rows", headerStyle = hs1)

hs2 <- createStyle(fontColour = "#ffffff", fgFill = "#4F80BD",
                   halign = "center", valign = "center", textDecoration = "Bold",
                   border = "TopBottomLeftRight", borderColour = "#4F81BD")

writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
          startCol="O", startRow = 23, borders="columns", headerStyle = hs2)

## Save workbook
saveWorkbook(wb, "writeDataExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{writeData}}
}

