\name{loadWorkbook}
\alias{loadWorkbook}
\title{Load an exisiting .xlsx file}
\usage{
loadWorkbook(xlsxFile)
}
\arguments{
  \item{xlsxFile}{A path to an existing .xlsx file}
}
\value{
Workbook object.
}
\description{
loadWorkbook returns a workbook object conserving styles
and formatting of the original .xlsx file.
}
\examples{
## load existing workbook from package folder
wb <- loadWorkbook(xlsxFile = system.file("loadExample.xlsx", package= "openxlsx"))
sheets(wb)  #list worksheets
wb ## view object
## Add a worksheet
addWorksheet(wb, "A new worksheet")

## update data in Sales worksheet
df <- data.frame("Orders" = round(runif(10)*200),
                 "TotalSales" = runif(10)*400)
writeData(wb, "Sales", df, xy=c("A", 2))

## Save workbook
saveWorkbook(wb, "loadExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{removeWorksheet}}
}

