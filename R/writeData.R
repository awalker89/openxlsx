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
#' @param borderColour Colour of cell border.  A valid colour (belonging to colours() or a hex colour code.)
#' @param borderStyle Border line style
#' \itemize{
#'    \item{\bold{none}}{ No Border}
#'    \item{\bold{thin}}{ thin border}
#'    \item{\bold{medium}}{ medium border}
#'    \item{\bold{dashed}}{ dashed border}
#'    \item{\bold{dotted}}{ dotted border}
#'    \item{\bold{thick}}{ thick border}
#'    \item{\bold{double}}{ double line border}
#'    \item{\bold{hair}}{ Hairline border}
#'    \item{\bold{mediumDashed}}{ medium weight dashed border}
#'    \item{\bold{dashDot}}{ dash-dot border}
#'    \item{\bold{mediumDashDot}}{ medium weight dash-dot border}
#'    \item{\bold{dashDotDot}}{ dash-dot-dot border}
#'    \item{\bold{mediumDashDotDot}}{ medium weight dash-dot-dot border}
#'    \item{\bold{slantDashDot}}{ slanted dash-dot border}
#'   }
#' @param ...  Further arguments (for future use)
#' @seealso \code{\link{writeData}}
#' @export writeData
#' @rdname writeData
#' @examples
#' wb <- createWorkbook()
#'
#' ## Add worksheets
#' addWorksheet(wb, "Cars")
#' 
#' x <- mtcars[1:6,]
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' writeData(wb, "Cars", x, rowNames = TRUE, startCol = "O", startRow = 3, 
#'          borders="surrounding", borderColour = NULL) ## black border
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol = 2, startRow = 12, borders="columns")
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol="O", startRow = 12, borders="rows", borderColour = "#4F81BD")
#'
#' ## header styles
#' hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "Italic",
#'                    border = "Bottom", borderColour = "green")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE, startCol="B",
#'      startRow = 23, borders="rows", headerStyle = hs1, borderStyle = "dashed")
#' 
#' hs2 <- createStyle(fontColour = "#ffffff", fgFill = "#4F80BD",
#'                    halign = "center", valign = "center", textDecoration = "Bold",
#'                    border = "TopBottomLeftRight", borderColour = "#4F81BD")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
#'           startCol="O", startRow = 23, borders="columns", headerStyle = hs2)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "writeDataExample.xlsx", overwrite = TRUE)
#' 
#' \dontrun{
#'
#' ## inspired by xtable gallery
#' ##' http://cran.r-project.org/web/packages/xtable/vignettes/xtableGallery.pdf
#' 
#' ## Create a new workbook
#' wb <- createWorkbook()
#' data(tli, package = "xtable")
#' 
#' ## TEST 1 - data.frame
#' test.n <- "data.frame"
#' my.df <- tli[1:10, ]
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.df, borders = "n")
#' 
#' ## TEST 2 - matrix
#' test.n <- "matrix"
#' design.matrix <- model.matrix(~ sex * grade, data = my.df)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = design.matrix)
#' 
#' ## TEST 3 - aov
#' test.n <- "aov"
#' fm1 <- aov(tlimth ~ sex + ethnicty + grade + disadvg, data = tli)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm1)
#' 
#' ## TEST 4 - lm
#' test.n <- "lm"
#' fm2 <- lm(tlimth ~ sex*ethnicty, data = tli)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm2)
#' 
#' ## TEST 5 - anova 1 
#' test.n <- "anova"
#' my.anova <- anova(fm2)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.anova)
#' 
#' ## TEST 6 - anova 2
#' test.n <- "anova2"
#' fm2b <- lm(tlimth ~ ethnicty, data = tli)
#' my.anova2 <- anova(fm2b, fm2)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.anova2)
#' 
#' ## TEST 7 - GLM
#' test.n <- "glm"
#' fm3 <- glm(disadvg ~ ethnicty*grade, data = tli, family =
#'            binomial())
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm3)
#'
#' ## TEST 8 - simple table
#' test.n <- "table"
#' data(airquality)
#' airquality$OzoneG80 <- factor(airquality$Ozone > 80,
#'                               levels = c(FALSE, TRUE),
#'                               labels = c("Oz <= 80", "Oz > 80"))
#' airquality$Month <- factor(airquality$Month,
#'                            levels = 5:9,
#'                            labels = month.abb[5:9])
#' my.table <- with(airquality, table(OzoneG80,Month) )
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.table)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "classTests.xlsx",  overwrite = TRUE)
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
                      borderStyle = getOption("openxlsx.borderStyle", "thin"),
                      ...){
  
  ## All input conversions/validations
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
  
  borders <- match.arg(borders)
  if(length(borders) != 1) stop("borders argument must be length 1.")    
  
  ## borderColours validation
  borderColour <- validateColour(borderColour)
  borderStyle <- validateBorderStyle(borderStyle)[[1]]
  
  ## Have decided to not use S3 as too much code duplication with input checking/converting
  ## given that everything has to fit into a grid.
  
  clx <- class(x)
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
    
  }else{
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
  
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style  
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             rows = startRow,
             cols = 0:(nCol-1) + startCol,
             gridExpand = TRUE)
  
  ## draw borders
  if( "none" != borders )
    doBorders(borders = borders, wb = wb,
              sheet = sheet, startRow = startRow + colNames,
              startCol = startCol, nrow = nrow(x), ncol = ncol(x),
              borderColour = borderColour, borderStyle = borderStyle)
  
}




