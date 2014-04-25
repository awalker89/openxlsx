\name{removeColWidths}
\alias{removeColWidths}
\title{Remove custom column widths from a worksheet}
\usage{
removeColWidths(wb, sheet, cols)
}
\arguments{
  \item{wb}{A workbook object}

  \item{sheet}{A name or index of a worksheet}

  \item{cols}{Indices of colunss to remove custom width (if
  any) from.}
}
\description{
Remove custom column widths from a worksheet
}
\examples{
## Create a new workbook
wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))

## remove column widths in columns 1 to 10
removeColWidths(wb, 1, cols = 1:10)
saveWorkbook(wb, "removeColWidthsExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{setColWidths}}
}

