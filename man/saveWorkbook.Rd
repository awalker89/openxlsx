\name{saveWorkbook}
\alias{saveWorkbook}
\title{save Workbook to file}
\usage{
saveWorkbook(wb, file, overwrite = FALSE)
}
\arguments{
  \item{wb}{A Workbook object to write to file}

  \item{file}{A character string naming an xlsx file}

  \item{overwrite}{If TRUE, overwrite any existing file.}
}
\description{
save Workbook to file
}
\examples{
## Create a new workbook and add a worksheet
wb <- createWorkbook("Creator of workbook")
addWorksheet(wb, sheetName = "My first worksheet")

## Save workbook to working directory
saveWorkbook(wb, file = "saveWorkbookExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{createWorkbook}}

\code{\link{addWorksheet}}

\code{\link{loadWorkbook}}

\code{\link{writeData}}

\code{\link{writeDataTable}}
}

