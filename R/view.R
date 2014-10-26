#'
#' Alternative to View that uses openXL
#'
#' Alternative to \code{\link{View}} that invokes a spreasheet viewer on
#' openxlsx's \code{\link{writeData}} handled objects. 
#'
#' @param ... writeData handled objects
#' @examples \dontrun{
#' view(Indometh, iris)
#' }
#' @export
view <- function(...)  {
  wb <- createWorkbook()
  objList <- list(...)
  objNames <- lapply(as.list(match.call(expand.dots = TRUE))[-1],
                     as.character)
  mapply(addWorksheet,
         sheetName = objNames,
         MoreArgs = list(wb = wb))
  mapply(writeData,
         sheet = objNames,
         x = objList,
         MoreArgs = list(wb = wb)) 
  openXL(wb)
}







  
