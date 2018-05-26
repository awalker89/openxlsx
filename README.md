openxlsx
========
This [R](https://www.R-project.org/) package simplifies the
creation of `.xlsx` files by providing 
a high level interface to writing, styling and editing
worksheets. Through the use of
[`Rcpp`](https://CRAN.R-project.org/package=Rcpp), 
read/write times are comparable to the
[`xlsx`](https://CRAN.R-project.org/package=xlsx)
and
[`XLConnect`](https://CRAN.R-project.org/package=XLConnect)
packages with the added benefit of removing the dependency on
Java. 

## Installation

### Stable version
Current stable version is available on
[CRAN](https://CRAN.R-project.org/) via
```R
install.packages("openxlsx", dependencies = TRUE)
```

### Development version
```R
install.packages(c("Rcpp", "devtools"), dependencies = TRUE)
require(devtools)
install_github("awalker89/openxlsx")
```

## Bug/feature request
Please let me know which version of openxlsx you are using when posting bug reports.
```R
packageVersion("openxlsx")
```
Thanks, [here](https://github.com/awalker89/openxlsx/issues). 

## News
[Here](https://raw.githubusercontent.com/awalker89/openxlsx/master/NEWS). 

## Authors and Contributors
A list is automagically maintained
[here](https://github.com/awalker89/openxlsx/graphs/contributors). 