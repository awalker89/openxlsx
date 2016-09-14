openxlsx
========
This [R](http://www.r-project.org/) package simplifies the
creation of `.xlsx` files by providing
a high level interface to writing, styling and editing
worksheets. Through the use of
[`Rcpp`](http://cran.r-project.org/web/packages/Rcpp/),
read/write times are comparable to the
[`xlsx`](http://cran.r-project.org/web/packages/xlsx/index.html)
and
[`XLConnect`](http://cran.r-project.org/web/packages/XLConnect)
packages with the added benefit of removing the dependency on
Java.

## Installation

The openxlsx package requires a zip application to be available to R,
 such as the one that comes with Rtools, available [here](http://cran.r-project.org/bin/windows/Rtools/). (Windows only)

 If the command
 ```R
 shell("zip")
 ```
 returns
 ```R
'zip' is not recognized as an internal or external command, operable program or
 batch file.
 ```

 or similar.  Then;

 * Install Rtools from: http://cran.r-project.org/bin/windows/Rtools/ and modify
 the system PATH during installation.

 * If Rtools is installed, add the Rtools bin directory paths (default installation paths are
 c:\Rtools\bin and c:\Rtools\gcc-4.6.3\bin) to the system PATH variable.

### Stable version
Current stable version is available on
[CRAN](http://cran.r-project.org/) via
```R
install.packages("openxlsx", dependencies=TRUE)
```

### Development version
Development version can be installed via GitHub once Rtools (Windows only) has been setup with:

```R
install.packages(c("Rcpp", "devtools"), dependencies=TRUE)
require(devtools)
install_github("awalker89/openxlsx")
```

## Bug/feature request
Thanks, [here](https://github.com/awalker89/openxlsx/issues).

## News
[Here](https://raw.githubusercontent.com/awalker89/openxlsx/master/NEWS).

## Authors and Contributors
A list is automagically maintained
[here](https://github.com/awalker89/openxlsx/graphs/contributors).
