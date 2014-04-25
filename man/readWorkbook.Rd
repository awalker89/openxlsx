\name{readWorkbook}
\alias{readWorkbook}
\title{Read data from a worksheet into a data.frame}
\usage{
readWorkbook(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE)
}
\arguments{
  \item{xlsxFile}{A xlsx workbook}

  \item{sheet}{The name or index of the sheet to read data
  from}

  \item{startRow}{startRow is used to specify where to
  begin looking for data.  Empty rows will be skipped
  regardless of the value of startRow.}

  \item{colNames}{IF TRUE, will attempt to look for column
  names.}
}
\value{
data.frame
}
\description{
Read data from a worksheet into a data.frame
}
\details{
Creates a data.frame of all data in worksheet. If the first
row of data consists entirely of strings, these will be
used as column names.
}
\examples{
xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
df1 <- readWorkbook(xlsxFile = xlsxFile, sheet = 1)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{read.xlsx}}
}

