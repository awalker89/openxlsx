


#' @name writeData
#' @title Write an object to a worksheet
#' @author Alexander Walker
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written.
#' @param startCol A vector specifiying the starting columns(s) to write df
#' @param startRow A vector specifiying the starting row(s) to write df
#' @param xy An alternative to specifying startCol and startRow individually.
#'  A vector of the form c(startCol, startRow)
#' @param colNames If TRUE, column names of x are written.
#' @param rowNames If TRUE, data.frame row names of x are written.
#' @param headerStyle Custom style to apply to column names.
#' @param borders Either "none" (default), "surrounding",
#' "columns", "rows" or respective abbreviations.  If
#' "surrounding", a border is drawn around the data.  If "rows",
#' a surrounding border is drawn a border around each row. If
#' "columns", a surrounding border is drawn with a border between
#' each column.
#' @param borderColour Colour of cell border.  A valid colour (belonging to colours())
#' @param ...  Further arguments (for future use)
#' @seealso \code{\link{writeData}}
#' @export writeData
#' @rdname writeData
#' @examples
#' \dontrun{
#' ## inspired by xtable gallery
#' ## http://cran.r-project.org/web/packages/xtable/vignettes/xtableGallery.pdf
#' ## Create a new workbook and delete old file, if existing
#' wb <- createWorkbook()
#' my.file <- "test.xlsx"
#' unlink(my.file)
#' data(tli, package = "xtable")
#' ## TEST 1 - data.frame
#' test.n <- "data.frame"
#' my.df <- tli[1:10, ]
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.df, borders = "n")
#' ## TEST 2 - matrix
#' test.n <- "matrix"
#' design.matrix <- model.matrix(~ sex * grade, data = my.df)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = design.matrix)
#' ## TEST 3 - aov
#' test.n <- "aov"
#' fm1 <- aov(tlimth ~ sex + ethnicty + grade + disadvg, data = tli)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm1)
#' ## TEST 4 - lm
#' test.n <- "lm"
#' fm2 <- lm(tlimth ~ sex*ethnicty, data = tli)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm2)
#' ## TEST 5 - anova 1 
#' test.n <- "anova"
#' my.anova <- anova(fm2)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.anova)
#' ## TEST 6 - anova 2
#' test.n <- "anova2"
#' fm2b <- lm(tlimth ~ ethnicty, data = tli)
#' my.anova2 <- anova(fm2b, fm2)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.anova2)
#' ## TEST 7 - GLM
#' test.n <- "glm"
#' fm3 <- glm(disadvg ~ ethnicty*grade, data = tli, family =
#'            binomial())
#'            
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm3)
#' ## Save workbook
#' saveWorkbook(wb, my.file,  overwrite = TRUE)
#' }
writeData <- function(wb, 
                       sheet,
                       x,
                       startCol = 1,
                       startRow = 1, 
                       xy = NULL,
                       colNames = TRUE,
                       rowNames = FALSE,
                       headerStyle = NULL,
                       borders = c("none","surrounding","rows","columns"),
                       borderColour = getOption("openxlsx.borderColour", "black"),
                       ...){
  
  ## NOTE: writeData is the first piece of writeData
  ## NOTE: writeData.default is the second piece of writeData
  
  ## Call <- match.call(expand.dots = TRUE)
  ## Call$startCol
  
  ## Input validating
  if(!is.null(xy)){
    if(length(xy) != 2)
      stop("xy parameter must have length 2")
    startCol <-  xy[[1]]
    startRow <- xy[[2]]
  }
  
  ## convert startRow and startCol
  startCol <- convertFromExcelRef(startCol)
  startRow <- as.integer(startRow)

  if(!"Workbook" %in% class(wb)) stop("First argument must be a Workbook.")
  if(!is.logical(colNames)) stop("colNames must be a logical.")
  if(!is.logical(rowNames)) stop("rowNames must be a logical.")
  if(!is.null(headerStyle) & !"Style" %in% class(headerStyle)) stop("headerStyle must be a style object or NULL.")
      
  ## borderColours validation
  borderColour <- validateBorderColour(borderColour)
  
  ## Method dispatch
  UseMethod(generic = "writeData", object = x)
  
}


#' @method writeData default
#' @S3method writeData default
writeData.default <- function(wb, 
                               sheet,
                               x,
                               startCol = 1,
                               startRow = 1, 
                               xy = NULL,
                               colNames = TRUE,
                               rowNames = FALSE,
                               headerStyle = NULL,
                               borders = c("none","surrounding","rows","columns"),
                               borderColour = getOption("openxlsx.borderColour", "black"),
                               ...){
  
  borders <- match.arg(borders)
  
  if(!any(c("matrix", "data.frame", "data.table") %in% class(x))){
    x <- as.data.frame(x)
    colNames <- FALSE
    rowNames <- FALSE
  }
  
  if("matrix" %in% class(x))
    x <- as.data.frame(x)
  
  ##Coordinates for each section
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x)), x)
    names(x)[[1]] <- ""
  }
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## Get coordinated of each header row and data.frame cells
  headerCoords <- list("row" = startRow, "col" = 0:(nCol-1) + startCol)
  
  ## border style cases
  if( "none" != borders )
    doBorders(borders = borders, wb = wb,
              sheet = sheet, srow = startRow + colNames,
              scol = startCol, nrow = nrow(x), ncol = ncol(x),
              borderColour = borderColour)
  
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style  
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             headerCoords$row, headerCoords$col,
             gridExpand = TRUE)
  
}




writeData.array <- function(wb, 
                             sheet,
                             x,
                             startCol = 1,
                             startRow = 1, 
                             xy = NULL,
                             colNames = TRUE,
                             rowNames = FALSE,
                             headerStyle = NULL,
                             borders = c("none","surrounding","rows","columns"),
                             borderColour = getOption("openxlsx.borderColour", "black"),
                             ...){
  
  
  borders <- match.arg(borders)
  ## TODO: writeData.array
  stop("array in writeData : currently not supported")
}


#' @method writeData lm
#' @S3method writeData lm
writeData.lm <- function(wb, 
                          sheet,
                          x,
                          startCol = 1,
                          startRow = 1, 
                          xy = NULL,
                          colNames = TRUE,
                          rowNames = FALSE,
                          headerStyle = NULL,
                          borders = c("none","surrounding","rows","columns"), 
                          borderColour = getOption("openxlsx.borderColour", "black"),
                          ...){
  
  borders <- match.arg(borders)
  x <- as.data.frame(summary(x)[["coefficients"]])
  x <- cbind(data.frame("Variable" = rownames(x)), x)
  names(x)[1] <- ""
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## Get coordinated of each header row and data.frame cells
  headerCoords <- list("row" = startRow, "col" = 0:(nCol-1) + startCol)
  
  ## Write data and styling
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style and default
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             headerCoords$row, headerCoords$col, gridExpand = TRUE)
  
}


#' @method writeData aov
#' @S3method writeData aov
writeData.aov <- function(wb, 
                           sheet,
                           x,
                           startCol = 1,
                           startRow = 1, 
                           xy = NULL,
                           colNames = TRUE,
                           rowNames = FALSE,
                           headerStyle = NULL,
                           borders = c("none","surrounding","rows","columns"), 
                           borderColour = getOption("openxlsx.borderColour", "black"),
                           ...){

  borders <- match.arg(borders)
  
  x <- summary(x)
  x <- cbind(x[[1]])
  x <- cbind(data.frame("row name" = rownames(x)), x)
  names(x)[1] <- ""
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## Get coordinated of each header row and data.frame cells
  headerCoords <- list("row" = startRow, "col" = 0:(nCol-1) + startCol)
  
  ## Write data and styling
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style and default
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             headerCoords$row, headerCoords$col, gridExpand = TRUE)
  
}


#' @method writeData anova
#' @S3method writeData anova
writeData.anova <- function(wb, 
                             sheet,
                             x,
                             startCol = 1,
                             startRow = 1, 
                             xy = NULL,
                             colNames = TRUE,
                             rowNames = FALSE,
                             headerStyle = NULL,
                             borders = c("none","surrounding","rows","columns"), 
                             borderColour = getOption("openxlsx.borderColour", "black"),
                             ...){
  borders <- match.arg(borders)
  
  x <- cbind(x)
  x <- cbind(data.frame("row name" = rownames(x)), x)
  names(x)[1] <- ""
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## Get coordinated of each header row and data.frame cells
  headerCoords <- list("row" = startRow, "col" = 0:(nCol-1) + startCol)
  
  ## Write data and styling
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style and default
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             headerCoords$row, headerCoords$col, gridExpand = TRUE)
  
}


#' @method writeData glm
#' @S3method writeData glm
writeData.glm <- function(wb, 
                           sheet,
                           x,
                           startCol = 1,
                           startRow = 1, 
                           xy = NULL,
                           colNames = TRUE,
                           rowNames = FALSE,
                           headerStyle = NULL,
                           borders = c("none","surrounding","rows","columns"), 
                           borderColour = getOption("openxlsx.borderColour", "black"),
                           ...){
  
  borders <- match.arg(borders)
  x <- as.data.frame(summary(x)[["coefficients"]])
  x <- cbind(data.frame("row name" = rownames(x)), x)
  names(x)[1] <- ""
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## Get coordinated of each header row and data.frame cells
  headerCoords <- list("row" = startRow, "col" = 0:(nCol-1) + startCol)
  
  ## Write data and styling
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style and default
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             headerCoords$row, headerCoords$col, gridExpand = TRUE)
  
}
