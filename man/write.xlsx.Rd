\name{write.xlsx}
\alias{write.xlsx}
\title{write directly to an xlsx file}
\usage{
write.xlsx(x, file, ...)
}
\arguments{
  \item{x}{data to write to file}

  \item{file}{xlsx file name}

  \item{...}{optional parameters to pass to functions:
  \itemize{ \item{createWorkbook} \item{addWorksheet}
  \item{writeData} \item{saveWorkbook} }

  see detail}
}
\description{
write directly to an xlsx file
}
\details{
Optional parameters are: \itemize{ \item{\bold{creator}}{ A
string specifying the workbook author}
\item{\bold{sheetName}}{ Name of the worksheet}
\item{\bold{gridLines}}{ A logical. If FALSE, the worksheet
grid lines will be hidden.} \item{\bold{startCol}}{ A
vector specifiying the starting columns(s) to write df}
\item{\bold{startRow}}{ A vector specifiying the starting
row(s) to write df} \item{\bold{xy}}{ An alternative to
specifying startCol and startRow individually.  A vector of
the form c(startCol, startRow)} \item{\bold{colNames or
col.names}}{ If TRUE, column names of x are written.}
\item{\bold{rowNames or row.names}}{ If TRUE, row names of
x are written.} \item{\bold{headerStyle}}{ Custom style to
apply to column names.} \item{\bold{borders}}{ Either
"surrounding", "columns" or "rows" or NULL.  If
"surrounding", a border is drawn around the data.  If
"rows", a surrounding border is drawn a border around each
row. If "columns", a surrounding border is drawn with a
border between each column.} \item{\bold{borderColour}}{
Colour of cell border.  A valid hex colour beginning with
"#".  NULL will set borders to black.} }

columns of x with class Date, POSIXct of POSIXt are
automatically styled as dates.
}
\author{
Alexander Walker
}
\seealso{
\code{\link{addWorksheet}}

\code{\link{writeData}}
}

