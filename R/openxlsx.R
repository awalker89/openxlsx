#' xlsx reading, writing and editing.
#'
#' openxlsx simplifies the the process of writing and styling Excel xlsx files from R
#' and removes the dependency on Java.
#' 
#' @name openxlsx
#' @docType package
#' @useDynLib openxlsx, .registration=TRUE
#' @import grDevices
#' @import stats
#' @import utils
#' @importFrom Rcpp sourceCpp 
#' 
#' @seealso
#' \itemize{
#'    \item{\code{vignette("Introduction", package = "openxlsx")}}
#'    \item{\code{vignette("formatting", package = "openxlsx")}}
#'    \item{\code{\link{writeData}}}
#'    \item{\code{\link{writeDataTable}}}
#'    \item{\code{\link{write.xlsx}}}
#'    \item{\code{\link{read.xlsx}}}
#'   } 
#' for examples
#' 
#' @details
#' The openxlsx package uses global options to simplfy formatting:
#' 
#' \itemize{
#'    \item{\code{options("openxlsx.borderColour" = "black")}}
#'    \item{\code{options("openxlsx.borderStyle" = "thin")}}
#'    \item{\code{options("openxlsx.dateFormat" = "mm/dd/yyyy")}}
#'    \item{\code{options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")}}
#'    \item{\code{options("openxlsx.numFmt" = NULL)}}
#'    \item{\code{options("openxlsx.paperSize" = 9)}} ## A4
#'    \item{\code{options("openxlsx.orientation" = "portrait")}} ## page orientation
#' } 
#'  See the Formatting vignette for examples. 
#' 
#' 
#' 
#' 
#' Additional options
#' 
#' \itemize{
#' \item{\code{options("openxlsx.zipFlags" = "-9")}} ## set max zip compression level default is "-1"
#' } 
#' 
NULL