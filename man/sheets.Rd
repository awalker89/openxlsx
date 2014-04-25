\name{sheets}
\alias{sheets}
\title{Returns names of worksheets.}
\usage{
sheets(wb)
}
\arguments{
  \item{wb}{A workbook object}
}
\value{
Name of worksheet(s) for a given index
}
\description{
Returns names of worksheets.
}
\examples{
## Create a new workbook
wb <- createWorkbook()

## Add some worksheets
addWorksheet(wb, "Worksheet Name")
addWorksheet(wb, "This is worksheet 2")
addWorksheet(wb, "The third worksheet")

## Return names of sheets, can not be used for assignment.
sheets(wb)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{renameWorksheet}}
}

