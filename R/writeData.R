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
#' @param name If not NULL, a named region is defined.
#' @param sep Only applies to list columns. The seperator used to collapse list columns to a character vector e.g. sapply(x$list_column, paste, collapse = sep).
#' @seealso \code{\link{writeDataTable}}
#' @export writeData
#' @details Formulae written using writeFormula to a Workbook object will not get picked up by read.xlsx().
#' This is because only the formula is written and left to Excel to evaluate the formula when the file is opened in Excel.
#' @rdname writeData
#' @examples
#' 
#' ## See formatting vignette for further examples. 
#' 
#' ## Options for default styling (These are the defaults)
#' options("openxlsx.borderColour" = "black")
#' options("openxlsx.borderStyle" = "thin")
#' options("openxlsx.dateFormat" = "mm/dd/yyyy")
#' options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
#' options("openxlsx.numFmt" = NULL)
#' 
#' ## Change the default border colour to #4F81BD
#' options("openxlsx.borderColour" = "#4F81BD")
#' 
#' 
#' #####################################################################################
#' ## Create Workbook object and add worksheets
#' wb <- createWorkbook()
#' 
#' ## Add worksheets
#' addWorksheet(wb, "Cars")
#' addWorksheet(wb, "Formula")
#' 
#' 
#' x <- mtcars[1:6,]
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' 
#' #####################################################################################
#' ## Bordering
#' 
#' writeData(wb, "Cars", x, rowNames = TRUE, startCol = "O", startRow = 3, 
#'          borders="surrounding", borderColour = "black") ## black border
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol = 2, startRow = 12, borders="columns")
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol="O", startRow = 12, borders="rows")
#'
#'
#' #####################################################################################
#' ## Header Styles
#' 
#' hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "italic",
#'                    border = "Bottom")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE, startCol="B",
#'      startRow = 23, borders="rows", headerStyle = hs1, borderStyle = "dashed")
#' 
#' 
#' hs2 <- createStyle(fontColour = "#ffffff", fgFill = "#4F80BD",
#'                    halign = "center", valign = "center", textDecoration = "bold",
#'                    border = "TopBottomLeftRight")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
#'           startCol="O", startRow = 23, borders="columns", headerStyle = hs2)
#' 
#' 
#' 
#' 
#' ##################################################################################### 
#' ## Hyperlinks
#' ## - vectors/columns with class 'hyperlink' are written as hyperlinks'
#' 
#' v <- rep("http://cran.r-project.org/", 4)
#' names(v) <- paste("Hyperlink", 1:4) # Optional: names will be used as display text
#' class(v) <- 'hyperlink'
#' writeData(wb, "Cars", x = v, xy = c("B", 32))
#'
#'
#' ##################################################################################### 
#' ## Formulas
#' ## - vectors/columns with class 'formula' are written as formulas'
#' 
#' df <- data.frame(x=1:3, y = 1:3,
#'                  z = paste(paste0("A", 1:3+1L), paste0("B", 1:3+1L), sep = " + "),
#'                  stringsAsFactors = FALSE)
#' 
#' class(df$z) <- c(class(df$z), "formula")
#' 
#' writeData(wb, sheet = "Formula", x = df)
#' 
#' 
#' ##################################################################################### 
#' ## Save workbook
#' ## Open in excel without saving file: openXL(wb)
#' 
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
                      keepNA = FALSE,
                      name = NULL,
                      sep = ", "){
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- getOption("scipen")
  options("scipen" = 200)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  
  exDigits <- getOption("digits")
  options("digits" = 22)
  on.exit(options("digits" = exDigits), add = TRUE)

  
    
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
  if((!is.character(sep)) | (length(sep) != 1)) stop("sep must be a character vector of length 1")
  
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
  
  ## special case - formula
  if("formula" %in% class(x)){
    x <- data.frame("X" = x, stringsAsFactors = FALSE)
    class(x[[1]]) <- "formula"
    colNames = FALSE
  }

  ## named region
  if(!is.null(name)){ ## validate name
    ex_names <- regmatches(wb$workbook$definedNames, regexpr('(?<=name=")[^"]+', wb$workbook$definedNames, perl = TRUE))
    ex_names <- replaceXMLEntities(ex_names)
    
    if(name %in% ex_names){
      stop(sprintf("Named region with name '%s' already exists!", name))
    }else if(grepl("[^A-Z0-9_\\.]", name[1], ignore.case = TRUE)){
      stop("Invalid characters in name")
    }else if(grepl('^[A-Z]{1,3}[0-9]+$', name)){
      stop("name cannot look like a cell reference.")
    }
  }
  
  if(is.vector(x) | is.factor(x) | inherits(x, "Date"))
    colNames <- FALSE ## this will go to coerce.default and rowNames will be ignored 
  
  ## Coerce to data.frame
  x <- openxlsxCoerce(x, rowNames)
    
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## If no rows and not writing column names return as nothing to write
  if(nRow == 0 & !colNames)
    return(invisible(0))
    
  colClasses <- lapply(x, function(x) tolower(class(x)))
  colClasss2 <- colClasses
  colClasss2[sapply(colClasses, function(x) "formula" %in% x) & sapply(colClasses, function(x) "hyperlink" %in% x)] <- "formula"
    
  
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
               colClasses = colClasss2,
               hlinkNames = hlinkNames,
               keepNA = keepNA,
               list_sep = sep)
  
  ## header style  
  if("Style" %in% class(headerStyle) & colNames)
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             rows = startRow,
             cols = 0:(nCol-1) + startCol,
             gridExpand = TRUE)
  
  ## If we don't have any rows to write return
  if(nRow == 0)
    return(invisible(0))
  
  ## named region
  if(!is.null(name)){
    
    ref1 <- paste0("$", .Call("openxlsx_convert2ExcelRef", startCol, LETTERS), "$", startRow)
    ref2 <- paste0("$", .Call("openxlsx_convert2ExcelRef", startCol + nCol - 1L, LETTERS), "$", startRow + nRow - 1L + colNames)
    wb$createNamedRegion(ref1 = ref1, ref2 = ref2, name = name, sheet = names(wb$worksheets)[sheet])
    
  }
  
  
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












#' @name writeFormula
#' @title Write a character vector as an Excel Formula
#' @author Alexander Walker
#' @description Write a a character vector containing Excel formula to a worksheet
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A character vector.
#' @param startCol A vector specifiying the starting column to write to.
#' @param startRow A vector specifiying the starting row to write to.
#' @param xy An alternative to specifying \code{startCol} and
#' \code{startRow} individually.  A vector of the form
#' \code{c(startCol, startRow)}.
#' @seealso \code{\link{writeData}}
#' @export writeFormula
#' @rdname writeFormula
#' @examples
#' 
#' ## There are 3 ways to write a formula
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' writeData(wb, "Sheet 1", x = iris)
#' 
#' ## SEE int2col() to convert int to Excel column label
#' 
#' ## 1. -  As a character vector using writeFormula
#' 
#' v <- c("SUM(A2:A151)", "AVERAGE(B2:B151)") ## skip header row
#' writeFormula(wb, sheet = 1, x = v, startCol = 10, startRow = 2)
#' writeFormula(wb, 1, x = "A2 + B2", startCol = 10, startRow = 10)
#' 
#' 
#' ## 2. - As a data.frame column with class "formula" using writeData
#' 
#' df <- data.frame(x=1:3,
#'                  y = 1:3,
#'                  z = paste(paste0("A", 1:3+1L), paste0("B", 1:3+1L), sep = " + "),
#'                  z2 = sprintf("ADDRESS(1,%s)", 1:3),
#'                  stringsAsFactors = FALSE)
#' 
#' class(df$z) <- c(class(df$z), "formula")
#' class(df$z2) <- c(class(df$z2), "formula")
#' 
#' addWorksheet(wb, "Sheet 2")
#' writeData(wb, sheet = 2, x = df)
#' 
#' 
#' 
#' ## 3. - As a vector with class "formula" using writeData
#' 
#' v2 <- c("SUM(A2:A4)", "AVERAGE(B2:B4)", "MEDIAN(C2:C4)")
#' class(v2) <- c(class(v2), "formula")
#' 
#' writeData(wb, sheet = 2, x = v2, startCol = 10, startRow = 2)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "writeFormulaExample.xlsx", overwrite = TRUE)
#' 
writeFormula <- function(wb, 
                         sheet,
                         x,
                         startCol = 1,
                         startRow = 1, 
                         xy = NULL){

  if(!"character" %in% class(x))
    stop("x must be a character vector.")
  
  dfx <- data.frame("X" = x, stringsAsFactors = FALSE)
  class(dfx$X) <- c("character", "formula")
  
  if(grepl("^(=|)HYPERLINK\\(", x, ignore.case = TRUE))
    class(dfx$X) <- c("character", "formula", "hyperlink")
    
    
  
  writeData(wb = wb,
            sheet = sheet,
            x = dfx,
            startCol = startCol,
            startRow = startRow, 
            xy = xy,
            colNames = FALSE,
            rowNames = FALSE)
  
  
  invisible(0)
  
}

