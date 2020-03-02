openxlsx
========

[![Project Status: Active – The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![Build Status](https://travis-ci.org/awalker89/openxlsx.svg?branch=master)](https://travis-ci.org/awalker89/openxlsx)
[![Coverage Status](https://codecov.io/github/awalker89/openxlsx/coverage.svg?branch=master)](https://codecov.io/github/awalker89/openxlsx?branch=master)
[![CRAN_Status_Badge](http://www.r-pkg.org/badges/version/openxlsx)](http://cran.r-project.org/package=openxlsx)
[![CRAN RStudio mirror downloads](http://cranlogs.r-pkg.org/badges/openxlsx)](http://cran.r-project.org/web/packages/openxlsx/index.html)

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