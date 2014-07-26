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
#' @seealso See 
#' \itemize{
#'    \item{\code{vignette("Introduction", package = "openxlsx")}}
#'    \item{\code{vignette("formatting", package = "openxlsx")}}
#'   } 
#' for examples
#' 
#' @details
#' The openxlsx package uses global options to simplfy formatting:
#' 
#' \itemize{
#'    \item{\code{options("openxlsx.borderColour" = "black")}}
#'    \item{\code{options("openxlsx.borderStyle" = "thin")}}
#'    \item{\code{ooptions("openxlsx.dateFormat" = "mm/dd/yyyy")}}
#'    \item{\code{options("openxlsx.dateFormat" = "mm/dd/yyyy")}}
#'   } 
#' 
#' 
NULL