#' @name writeData
#' @title Write an object to a worksheet
#' @author Alexander Walker
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written. For classes supported look at the examples.
#' @param startCol A vector specifiying the starting column to write to.
#' @param startRow A vector specifiying the starting row to write to.
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
#' @param withFilter If \code{TRUE}, add filters to the column name row. NOTE can only have one filter per worksheet. 
#' @param keepNA If \code{TRUE}, NA values are converted to #N/A in Excel else NA cells will be empty.
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
                      withFilter = FALSE,
                      keepNA = FALSE){
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- getOption("scipen")
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
    
  ## special case - vector of hyperlinks
  hlinkNames <- NULL
  if("hyperlink" %in% class(x)){
    hlinkNames <- names(x)
    colNames = FALSE
  }
  
  if(is.vector(x) | is.factor(x))
    colNames <- FALSE ## this will go to coerce.default and rowNames will be ignored 
  
  ## Coerce to data.frame
  x <- openxlsxCoerce(x, rowNames)
    
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## If no rows and not writing column names return as nothing to write
  if(nRow == 0 & !colNames)
    return(invisible(0))
    
  colClasses <- lapply(x, function(x) tolower(class(x)))
  

  
  sheetX <- wb$validateSheet(sheet)
  if(wb$isChartSheet[[sheetX]]){
    stop("Cannot write to chart sheet.")
    return(NULL)
  }
    
  ## write autoFilter, can only have a single filter per worksheet
  if(withFilter)
    wb$worksheets[[sheetX]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', paste(getCellRefs(data.frame("x" = c(startRow, startRow), "y" = c(startCol, startCol + nCol - 1L))), collapse = ":"))
  
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow,
               colClasses = colClasses,
               hlinkNames = hlinkNames,
               keepNA = keepNA)
  
  ## header style  
  if("Style" %in% class(headerStyle) & colNames)
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             rows = startRow,
             cols = 0:(nCol-1) + startCol,
             gridExpand = TRUE)
  
  ## If we don't have any rows to write return
  if(nRow == 0)
    return(invisible(0))
  
  
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
