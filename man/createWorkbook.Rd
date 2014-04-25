\name{createWorkbook}
\alias{createWorkbook}
\title{Create a new Workbook object}
\usage{
createWorkbook(creator = Sys.getenv("USERNAME"))
}
\arguments{
  \item{creator}{Creator of the workbook (your name).
  Defaults to login username}
}
\value{
Workbook object
}
\description{
Create a new Workbook object
}
\examples{
## Create a new workbook
wb <- createWorkbook("Alex")

## Save workbook to working directory
saveWorkbook(wb, file = "createWorkbookExample.xlsx", overwrite = TRUE)
}
\author{
Alexander Walker
}
\seealso{
\code{\link{loadWorkbook}}

\code{\link{saveWorkbook}}
}

