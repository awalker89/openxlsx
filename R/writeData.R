#' @name writeData
#' @title Write an object to a worksheet
#' @author Alexander Walker
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written. For classes supported look at the examples.
#' @param startCol A vector specifiying the starting columns(s) to write to.
#' @param startRow A vector specifiying the starting row(s) to write to.
#' @param xy An alternative to specifying \code{startCol} and
#' \code{startRow} individually.  A vector of the form
#' \code{c(startCol, startRow)}.
#' @param colNames If \code{TRUE}, column names of x are written.
#' @param rowNames If \code{TRUE}, data.frame row names of x are written.
#' @param headerStyle Custom style to apply to column names.
#' @param borders Either "\code{none}" (default), "\code{surrounding}",
#' "\code{columns}", "\code{rows}" or \emph{respective abbreviations}.  If
#' "\code{surrounding}", a border is drawn around the data.  If "\code{rows}",
#' a surrounding border is drawn with a border around each row. If
#' "\code{columns}", a surrounding border is drawn with a border between
#' each column. If "\code{all}" all cell borders are drawn.
#' @param borderColour Colour of cell border.  A valid colour (belonging to \code{colours()} or a hex colour code, eg see \href{http://www.colorpicker.com}{here}).
#' @param borderStyle Border line style
#' \itemize{
#'    \item{\bold{none}}{ no border}
#'    \item{\bold{thin}}{ thin border}
#'    \item{\bold{medium}}{ medium border}
#'    \item{\bold{dashed}}{ dashed border}
#'    \item{\bold{dotted}}{ dotted border}
#'    \item{\bold{thick}}{ thick border}
#'    \item{\bold{double}}{ double line border}
#'    \item{\bold{hair}}{ hairline border}
#'    \item{\bold{mediumDashed}}{ medium weight dashed border}
#'    \item{\bold{dashDot}}{ dash-dot border}
#'    \item{\bold{mediumDashDot}}{ medium weight dash-dot border}
#'    \item{\bold{dashDotDot}}{ dash-dot-dot border}
#'    \item{\bold{mediumDashDotDot}}{ medium weight dash-dot-dot border}
#'    \item{\bold{slantDashDot}}{ slanted dash-dot border}
#'   }
#' @param ...  Further arguments (for future use)
#' @seealso \code{\link{writeDataTable}}
#' @export writeData
#' @rdname writeData
#' @examples
#' 
#' ## See formatting vignette for further examples. 
#' 
#' wb <- createWorkbook()
#' options("openxlsx.borderStyle" = "thin")
#' options("openxlsx.borderColour" = "#4F81BD")
#' ## Add worksheets
#' addWorksheet(wb, "Cars")
#' 
#' 
#' x <- mtcars[1:6,]
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' writeData(wb, "Cars", x, rowNames = TRUE, startCol = "O", startRow = 3, 
#'          borders="surrounding", borderColour = "black") ## black border
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol = 2, startRow = 12, borders="columns")
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol="O", startRow = 12, borders="rows")
#'
#' ## header styles
#' hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "Italic",
#'                    border = "Bottom")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE, startCol="B",
#'      startRow = 23, borders="rows", headerStyle = hs1, borderStyle = "dashed")
#' 
#' hs2 <- createStyle(fontColour = "#ffffff", fgFill = "#4F80BD",
#'                    halign = "center", valign = "center", textDecoration = "Bold",
#'                    border = "TopBottomLeftRight")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
#'           startCol="O", startRow = 23, borders="columns", headerStyle = hs2)
#' 
#' ## Write a hyperlink vector
#' v <- rep("http://cran.r-project.org/", 4)
#' names(v) <- paste("Hyperlink", 1:4) # Names will be used as display text
#' class(v) <- 'hyperlink'
#' writeData(wb, "Cars", x = v, xy = c("B", 32))

#' ## Save workbook
#' saveWorkbook(wb, "writeDataExample.xlsx", overwrite = TRUE)
writeData <- function(wb, 
                      sheet,
                      x,
                      startCol = 1,
                      startRow = 1, 
                      xy = NULL,
                      colNames = TRUE,
                      rowNames = FALSE,
                      headerStyle = NULL,
                      borders = c("none","surrounding","rows","columns", "all"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      ...){
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- options("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  if(is.null(x))
    return(invisible(0))
  
  ## All input conversions/validations
  if(!is.null(xy)){
    if(length(xy) != 2)
      stop("xy parameter must have length 2")
    startCol <-  xy[[1]]
    startRow <- xy[[2]]
  }
  
  ## convert startRow and startCol
  if(!is.numeric(startCol))
    startCol <- convertFromExcelRef(startCol)
  startRow <- as.integer(startRow)
  
  if(!"Workbook" %in% class(wb)) stop("First argument must be a Workbook.")
  if(!is.logical(colNames)) stop("colNames must be a logical.")
  if(!is.logical(rowNames)) stop("rowNames must be a logical.")
  if(!is.null(headerStyle) & !"Style" %in% class(headerStyle)) stop("headerStyle must be a style object or NULL.")
  
  borders <- match.arg(borders)
  if(length(borders) != 1) stop("borders argument must be length 1.")    
  
  ## borderColours validation
  borderColour <- validateColour(borderColour, "Invalid border colour")
  borderStyle <- validateBorderStyle(borderStyle)[[1]]
  
  ## Have decided to not use S3 as too much code duplication with input checking/converting
  ## given that everything has to fit into a grid.
  
  clx <- class(x)
  hlinkNames <- NULL
  if(any(c("data.frame", "data.table") %in% clx)){
    ## Do nothing
    
  }else if("matrix" %in% clx){
    x <- as.data.frame(x, stringsAsFactors = FALSE)
    
  }else if("array" %in% clx){
    stop("array in writeData : currently not supported")
    
  }else if("aov" %in% clx){
    
    x <- summary(x)
    x <- cbind(x[[1]])
    x <- cbind(data.frame("row name" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE 
    
  }else if("lm" %in% clx){
    
    x <- as.data.frame(summary(x)[["coefficients"]])
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("anova" %in% clx){
    
    x <- cbind(x)
    x <- cbind(data.frame("row name" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("glm" %in% clx){
    
    x <- as.data.frame(summary(x)[["coefficients"]])
    x <- cbind(data.frame("row name" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("table" %in% clx){
    
    x <- as.data.frame(unclass(x))
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("prcomp" %in% clx){
    
    x <- as.data.frame(x$rotation)
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("summary.prcomp" %in% clx){
    
    x <- as.data.frame(x$importance)
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE

  }else if("survdiff" %in% clx){

    ## like print.survdiff with some ideas from the ascii package
    if (length(x$n) == 1) {
      z <- sign(x$exp - x$obs) * sqrt(x$chisq)
      temp <- c(x$obs, x$exp, z, 1 - pchisq(x$chisq,  1))
      names(temp) <- c("Observed", "Expected", "Z", "p")
      x <- as.data.frame(t(temp))
    }
    else {                              
      if (is.matrix(x$obs)) {
        otmp <- apply(x$obs, 1, sum)
        etmp <- apply(x$exp, 1, sum)
      }
      else {
        otmp <- x$obs
        etmp <- x$exp
      }
      chisq <- c(x$chisq, rep(NA, length(x$n) - 1))
      df <- c((sum(1 * (etmp > 0))) - 1, rep(NA, length(x$n) - 1))
      p <- c(1 - pchisq(x$chisq, df[!is.na(df)]), rep(NA, length(x$n) - 1))
      temp <- cbind( x$n, otmp, etmp, 
                    ((otmp - etmp)^2)/etmp, ((otmp - etmp)^2)/diag(x$var),
                    chisq, df, p)
      colnames(temp) <- c("N", "Observed", "Expected", "(O-E)^2/E", "(O-E)^2/V",
                          "Chisq", "df","p")
      temp <- as.data.frame(temp, checknames = FALSE)
      x <- cbind("Group" = names(x$n), temp)
      names(x)[1] <- ""
    }
    rowNames <- FALSE

  }else if("coxph" %in% clx){
    ## sligthly modified print.coxph
    coef <- x$coefficients
    se <- sqrt(diag(x$var))
    if (is.null(coef) | is.null(se)) 
      stop("Input is not valid")
    if (is.null(x$naive.var)) {
      tmp <- cbind(coef, exp(coef), se, coef/se, pchisq((coef/se)^2, 1))
      colnames(tmp) <- c("coef", "exp(coef)", "se(coef)", "z", "p")
    }
    else {
      nse <- sqrt(diag(x$naive.var))
      tmp <- cbind(coef, exp(coef), nse, se, coef/se, pchisq((coef/se)^2, 1))
      colnames(tmp) <- c("coef", "exp(coef)", "se(coef)", "robust se", "z", "p")
    }
    x <- cbind("Variable" = names(coef), as.data.frame(tmp, checknames = FALSE))
    names(x)[1] <- ""
    rowNames <- FALSE

  }else if("summary.coxph" %in% clx){

    coef <- x$coefficients
    ci <- x$conf.int
    nvars <- nrow(coef)
    tmp <- cbind(coef[, - ncol(coef), drop=FALSE],          #p later
                 ci[, (ncol(ci) - 1):ncol(ci), drop=FALSE], #confint
                 coef[, ncol(coef), drop=FALSE])            #p.value
    x <- as.data.frame(tmp, checknames = FALSE)
    rowNames <- TRUE
    
  }else if("cox.zph" %in% clx){

    x <- as.data.frame(x$table)
    rowNames <- TRUE
    
  }else{
    
    if('hyperlink' %in% tolower(class(x))){
      hlinkNames <- names(x)
      class(x) <- c("character", "hyperlink") 
    }
    x <- as.data.frame(x, stringsAsFactors = FALSE)
    colNames <- FALSE
    rowNames <- FALSE
  }
  
  ## cbind rownames to x
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x)), x)
    names(x)[[1]] <- ""
  }
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  colClasses <- lapply(x, function(x) tolower(class(x)))
  sheet <- wb$validateSheet(sheet)

  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow,
               colClasses = colClasses,
               hlinkNames = hlinkNames)
  
  ## header style  
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             rows = startRow,
             cols = 0:(nCol-1) + startCol,
             gridExpand = TRUE)
  
  
  ## hyperlink style, if no borders
  if(borders == "none"){
    
    invisible(classStyles(wb, sheet = sheet, startRow = startRow, startCol = startCol, colNames = colNames, nRow = nrow(x), colClasses = colClasses))
    
  }else if(borders == "surrounding"){
    
    wb$surroundingBorders(colClasses,
                          sheet = sheet,
                          startRow = startRow + colNames,
                          startCol = startCol,
                          nRow = nRow, nCol = nCol,
                          borderColour = list("rgb" = borderColour),
                          borderStyle = borderStyle)
    
  }else if(borders == "rows"){
    
    wb$rowBorders(colClasses,
                  sheet = sheet,
                  startRow = startRow + colNames,
                  startCol = startCol,
                  nRow = nRow, nCol = nCol,
                  borderColour = list("rgb" = borderColour),
                  borderStyle = borderStyle)
    
  }else if(borders == "columns"){
    
    wb$columnBorders(colClasses,
                     sheet = sheet,
                     startRow = startRow + colNames,
                     startCol = startCol,
                     nRow = nRow, nCol = nCol,
                     borderColour = list("rgb" = borderColour),
                     borderStyle = borderStyle)
    
    
  }else if(borders == "all"){
    
    wb$allBorders(colClasses,
                  sheet = sheet,
                  startRow = startRow + colNames,
                  startCol = startCol,
                  nRow = nRow, nCol = nCol,
                  borderColour = list("rgb" = borderColour),
                  borderStyle = borderStyle)
    
    
  }
  
  invisible(0)
  
}




