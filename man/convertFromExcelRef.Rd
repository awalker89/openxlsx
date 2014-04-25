\name{convertFromExcelRef}
\alias{convertFromExcelRef}
\title{Convert excel column name to integer index}
\usage{
convertFromExcelRef(col)
}
\arguments{
  \item{col}{An excel column reference}
}
\description{
Convert excel column name to integer index
}
\examples{
convertFromExcelRef("DOG")
convertFromExcelRef("COW")

## numbers will be removed
convertFromExcelRef("R22")
}
\author{
Alexander Walker
}

