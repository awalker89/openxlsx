\name{read.xlsx}
\alias{read.xlsx}
\title{Read data from a worksheet into a data.frame}
\usage{
read.xlsx(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE,
  skipEmptyRows = TRUE)
}
\arguments{
  \item{xlsxFile}{An xlsx file}

  \item{sheet}{The name or index of the sheet to read data}

  \item{startRow}{first row to begin looking for data.
  Empty rows before any data is found are skipped.
  regardless of the value of startRow.}

  \item{colNames}{If TRUE, first row of data will be used
  as column names.}

  \item{skipEmptyRows}{If TRUE, empty rows are skipped else
  empty rows after the first row containing data will
  return a row of NAs}
}
\value{
data.frame
}
\description{
Read data from a worksheet into a data.frame
}
\details{
Creates a data.frame of all data in worksheet.
}
\examples{
xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
df1 <- read.xlsx(xlsxFile = xlsxFile, sheet = 1, startRow=1, skipEmptyRows=FALSE, colNames=TRUE)
sapply(df1, class)

df2 <- read.xlsx(xlsxFile = xlsxFile, sheet = 3, startRow=1, skipEmptyRows=TRUE, colNames=TRUE)
df2$Date <- convertToDate(df2$Date)
sapply(df2, class)
head(df2)

df3 <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, startRow=1, skipEmptyRows=TRUE, colNames=TRUE)
df3$Symbol
}
\author{
Alexander Walker
}
\seealso{
\code{\link{readWorkbook}}
}

