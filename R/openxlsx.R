#' xlsx reading, writing and editing.
#'
#' openxlsx simplifies the the process of writing and styling xlsx files from R
#' and removes the dependency on Java.
#' 
#' @name openxlsx
#' @docType package
#' @useDynLib openxlsx
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
#' } 
#'  See the Formatting vignette for examples. 
#' 
NULL