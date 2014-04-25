\name{renameWorksheet}
\alias{renameWorksheet}
\title{Rename an exisiting worksheet}
\usage{
renameWorksheet(wb, sheet, newName)
}
\arguments{
  \item{wb}{A Workbook object containing a worksheet}

  \item{sheet}{The name or index of the worksheet to
  rename}

  \item{newName}{The new name of the worksheet. No longer
  than 28 chars.}
}
\description{
Rename an exisiting worksheet
}
\examples{
## Create a new workbook
wb <- createWorkbook("CREATOR")

## Add 3 worksheets
addWorksheet(wb, "Worksheet Name")
addWorksheet(wb, "This is worksheet 2")
addWorksheet(wb, "Not the best name")

## Rename worksheet 1
renameWorksheet(wb, 1, "New name for sheet 1")

## Rename worksheet 3
renameWorksheet(wb, "Not the best name", "A better name")

## Save workbook
saveWorkbook(wb, "renameWorksheetExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}

