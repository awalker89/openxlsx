
#' @name createWorkbook
#' @title Create a new Workbook object
#' @param creator Creator of the workbook (your name). Defaults to login username
#' @author Alexander Walker
#' @return Workbook object
#' @export
#' @seealso \code{\link{loadWorkbook}}
#' @seealso \code{\link{saveWorkbook}}
#' @import methods
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Save workbook to working directory
#' saveWorkbook(wb, file = "createWorkbookExample", overwrite = TRUE)
createWorkbook <- function(creator = Sys.getenv("USERNAME")){
  
  if(class(creator) != "character")
    creator <- ""
  
  if(length(creator) > 1)
    creator <- creator[[1]]
  
  ## remove any illegal XML characters
  creator <- replaceIllegalCharacters(creator)
  
  invisible(Workbook$new(creator))
}


#' @name saveWorkbook
#' @title save Workbook to file
#' @author Alexander Walker
#' @param wb A Workbook object to write to file
#' @param file A character string naming an xlsx file
#' @param overwrite If TRUE, overwrite any existing file.
#' @seealso \code{\link{createWorkbook}}
#' @seealso \code{\link{addWorksheet}}
#' @seealso \code{\link{loadWorkbook}}
#' @seealso \code{\link{writeData}}
#' @seealso \code{\link{writeDataTable}}
#' @export
#' @examples
#' ## Create a new workbook and add a worksheet
#' wb <- createWorkbook("Creator of workbook")
#' addWorksheet(wb, sheetName = "My first worksheet")
#' 
#' ## Save workbook to working directory
#' saveWorkbook(wb, file = "saveWorkbookExample.xlsx", overwrite = TRUE) 
saveWorkbook <- function(wb, file, overwrite = FALSE){
  
  wd <- getwd()
  on.exit(setwd(wd), add = TRUE)
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- options("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!grepl("\\.xlsx", file))
    file <- paste0(file, ".xlsx")
  
  if(!is.logical(overwrite))
    overwrite = FALSE
  
  if(file.exists(file) & !overwrite)
    stop("File already exists!")
  
  tmp <- wb$saveWorkbook(quiet = TRUE)
  setwd(wd)
  
  file.copy(file.path(tmp$tmpDir, tmp$tmpFile), file, overwrite = overwrite)
  
  ## delete temporary dir
  unlink(tmp$tmpDir, force = TRUE, recursive = TRUE)
  
  invisible(1)
}


#' @name mergeCells
#' @title Merge cells within a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to merge
#' @param rows corresponding rows to merge
#' @details As merged region must be rectangular, only min and max of cols and rows are used.
#' @author Alexander Walker
#' @seealso \code{\link{removeCellMerge}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Add a worksheet
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#'
#' ## Merge cells: Row 2 column C to F (3:6)
#' mergeCells(wb, "Sheet 1", cols = 2, rows = 3:6)
#' 
#' ## Merge cells:Rows 10 to 20 columns A to J (1:10)
#' mergeCells(wb, 1, cols = 1:10, rows = 10:20)
#'
#' ## Intersecting merges
#' mergeCells(wb, 2, cols = 1:10, rows = 1)
#' mergeCells(wb, 2, cols = 5:10, rows = 2)
#' mergeCells(wb, 2, cols = c(1,10), rows = 12) ## equivalent to 1:10 as only min/max are used
#' #mergeCells(wb, 2, cols = 1, rows = c(1,10)) # Throws error because intersects existing merge
#' 
#' ## remove merged cells
#' removeCellMerge(wb, 2, cols = 1, rows = 1) # removes any intersecting merges
#' mergeCells(wb, 2, cols = 1, rows = 1:10) # Now this works
#'
#' ## Save workbook
#' saveWorkbook(wb, "mergeCellsExample.xlsx", overwrite = TRUE)
mergeCells <- function(wb, sheet, cols, rows){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)
  
  wb$mergeCells(sheet, startRow = min(rows), endRow = max(rows), startCol = min(cols), endCol = max(cols))
  
}


#' @name removeCellMerge
#' @title Create a new Workbook object
#' @description Unmerges any merged cells that intersect
#' with the region specified by, min(cols):max(cols) X min(rows):max(rows)
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols vector of column indices
#' @param rows vector of row indices
#' @author Alexander Walker
#' @export
#' @seealso \code{\link{mergeCells}}
removeCellMerge <- function(wb, sheet, cols, rows){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  cols <- convertFromExcelRef(cols)
  rows <- as.integer(rows)
  
  wb$removeCellMerge(sheet, startRow = min(rows), endRow = max(rows), startCol = min(cols), endCol = max(cols))
  
}


#' @name sheets
#' @title Returns names of worksheets.
#' @param wb A workbook object
#' @return Name of worksheet(s) for a given index
#' @author Alexander Walker
#' @seealso \code{\link{names}} to rename a worksheet in a Workbook
#' @details DEPRECATED. Use \code{\link{names}}
#' @export
#' @examples
#' 
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Add some worksheets
#' addWorksheet(wb, "Worksheet Name")
#' addWorksheet(wb, "This is worksheet 2")
#' addWorksheet(wb, "The third worksheet")
#' 
#' ## Return names of sheets, can not be used for assignment.
#' names(wb)
#' # openXL(wb)
#' 
#' names(wb) <- c("A", "B", "C")
#' names(wb)
#' # openXL(wb)
#' 
sheets <- function(wb){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  nms <- names(wb$worksheets)
  nms <- replaceXMLEntities(nms)
  
  return(nms)
}



#' @name addWorksheet
#' @title Add a worksheet to a workbook
#' @author Alexander Walker
#' @param wb A Workbook object to attach the new worksheet
#' @param sheetName A name for the new worksheet
#' @param gridLines A logical. If FALSE, the worksheet grid lines will be hidden.
#' @param tabColour Colour of the worksheet tab. A valid colour (belonging to colours()) or a valid hex colour beginning with "#"
#' @return XML tree
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook("Fred")
#' 
#' ## Add 3 worksheets
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2", gridLines = FALSE)
#' addWorksheet(wb, "Sheet 3", tabColour = "red")
#' addWorksheet(wb, "Sheet 4", gridLines = FALSE, tabColour = "#4F81BD")
#' 
#' ## Save workbook
#' saveWorkbook(wb, "addWorksheetExample.xlsx", overwrite = TRUE)
addWorksheet <- function(wb, sheetName, gridLines = TRUE, tabColour = NULL){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(sheetName %in% names(wb$worksheets))
    stop("A worksheet by that name already exists!")
  
  if(!is.logical(gridLines) | length(gridLines) > 1)
    stop("gridLines must be a logical of length 1.")
  
  if(nchar(sheetName) > 31)
    stop("sheetName too long! Max length is 31 characters.")
  
  if(!is.null(tabColour))
    tabColour <- validateColour(tabColour, "Invalid tabColour in addWorksheet.")
  
  if(!is.character(sheetName))
    sheetName <- as.character(sheetName)
  
  ## Invalid XML characters
  sheetName <- replaceIllegalCharacters(sheetName)
  
  invisible(wb$addWorksheet(sheetName, gridLines, tabColour))
} 


#' @name renameWorksheet
#' @title Rename an exisiting worksheet
#' @author Alexander Walker
#' @param wb A Workbook object containing a worksheet
#' @param sheet The name or index of the worksheet to rename
#' @param newName The new name of the worksheet. No longer than 31 chars.
#' @details DEPRECATED. Use \code{\link{names}}
#' @export
#' @examples
#' 
#' ## Create a new workbook
#' wb <- createWorkbook("CREATOR")
#' 
#' ## Add 3 worksheets
#' addWorksheet(wb, "Worksheet Name")
#' addWorksheet(wb, "This is worksheet 2")
#' addWorksheet(wb, "Not the best name")
#' 
#' #' ## rename all worksheets
#' names(wb) <- c("A", "B", "C")
#' 
#' 
#' ## Rename worksheet 1 & 3
#' renameWorksheet(wb, 1, "New name for sheet 1")
#' names(wb)[[1]] <- "New name for sheet 1"
#' names(wb)[[3]] <-  "A better name"
#' 
#' ## Save workbook
#' saveWorkbook(wb, "renameWorksheetExample.xlsx", overwrite = TRUE)
renameWorksheet <- function(wb, sheet, newName){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  invisible(wb$setSheetName(sheet, newName))
}


#' @name convertFromExcelRef
#' @author Alexander Walker
#' @title Convert excel column name to integer index
#' @param col An excel column reference
#' @export
#' @examples
#' convertFromExcelRef("DOG")
#' convertFromExcelRef("COW")
#' 
#' ## numbers will be removed
#' convertFromExcelRef("R22")
convertFromExcelRef <- function(col){
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- options("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  col <- toupper(col)
  charFlag <- grepl("[A-Z]", col)
  if(any(charFlag)){
    col[charFlag] <- gsub("[0-9]", "", col[charFlag])
    d <- lapply(strsplit(col[charFlag], split = ""), function(x) match(rev(x), LETTERS))
    col[charFlag] <- unlist(lapply(1:length(d), function(i) sum(d[[i]]*(26^(0:(length(d[[i]])-1L)))) ))
  }
  
  col[!charFlag] <- as.integer(col[!charFlag])
  
  return(as.integer(col))
}



#' @name createStyle 
#' @title Create a cell style
#' @description Create a new style to apply to worksheet cells
#' @author Alexander Walker
#' @seealso \code{\link{addStyle}}
#' @param fontName A name of a font. Note the font name is not validated. If fontName is NULL,
#' the workbook base font is used. (Defaults to Calibri)
#' @param fontColour Colour of text in cell.  A valid hex colour beginning with "#"
#' or one of colours(). If fontColour is NULL, the workbook base font colours is used.
#' (Defaults to black)
#' @param fontSize Font size. A numeric greater than 0.
#' If fontSize is NULL, the workbook base font size is used. (Defaults to 11)
#' @param numFmt Cell formatting
#' \itemize{
#'   \item{\bold{GENERAL}}
#'   \item{\bold{NUMBER}}
#'   \item{\bold{CURRENCY}}
#'   \item{\bold{ACCOUNTING}}
#'   \item{\bold{DATE}}
#'   \item{\bold{LONGDATE}}
#'   \item{\bold{TIME}}
#'   \item{\bold{PERCENTAGE}}
#'   \item{\bold{FRACTION}}
#'   \item{\bold{SCIENTIFIC}}
#'   \item{\bold{COMMA}{  for comma seperated thousands}}
#'   \item{For date/datetime styling a combination of d, m, y and punctuation marks}
#'   \item{For numeric rouding use "0.00" with the preferred number of deciaml places}
#' }
#' 
#' @param border Cell border 
#' (Any combination of "Top", "Bottom", "Left", "Right" in any order).
#' \itemize{
#'    \item{\bold{Top}}{ Top border}
#'    \item{\bold{Bottom}}{ Bottom border}
#'    \item{\bold{Left}}{ Left border}
#'    \item{\bold{Right}}{ Right border}
#'    \item{\bold{TopBottom}}{ Top and bottom border}
#'    \item{\bold{LeftRight}}{ Left and right border}
#'    \item{\bold{TopLeftRight}}{ Top and Left border}
#'    \item{\bold{TopBottomLeftRight}}{ All borders}
#'   }
#'   
#' @param borderColour  Colour of cell border.  
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"  
#' 
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
#'    
#' @param bgFill Cell background fill colour. 
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"   
#' @param fgFill Cell foreground fill colour.
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"   
#'   
#' @param halign
#' Horizontal alignment of cell contents
#' \itemize{
#'    \item{\bold{left}}{ Left horizontal align cell contents}
#'    \item{\bold{right}}{ Right horizontal align cell contents}
#'    \item{\bold{center}}{ Center horizontal align cell contents}
#'   }   
#'   
#' @param valign A name
#' Vertical alignment of cell contents
#' \itemize{
#'    \item{\bold{top}}{ Top vertical align cell contents}
#'    \item{\bold{center}}{ Center vertical align cell contents}
#'    \item{\bold{bottom}}{ Bottom vertical align cell contents}
#'   } 
#'   
#' @param textDecoration
#' Text styling.
#' \itemize{
#'    \item{\bold{bold}}{ Bold cell contents}
#'    \item{\bold{strikeout}}{ Strikeout cell contents}
#'    \item{\bold{italic}}{ Italicise cell contents}
#'    \item{\bold{underline}}{ Underline cell contents}
#'    \item{\bold{underline2}}{ Double underline cell contents}
#'   } 
#'   
#' @param wrapText Logical. If TRUE cell contents will wrap to fit in column.  
#' @param textRotation Rotation of text in degrees. Numeric in [0, 180].
#' @return A style object
#' @export
#' @examples
#' ## See package vignettes for further examples
#' 
#' ## Modify default values of border colour and border line style
#' options("openxlsx.borderColour" = "#4F80BD")
#' options("openxlsx.borderStyle" = "thin")
#' 
#' ## Size 18 Arial, Bold, left horz. aligned, fill colour #1A33CC, all borders,
#' style <- createStyle(fontSize = 18, fontName = "Arial",
#'   textDecoration = "bold", halign = "left", fgFill = "#1A33CC", border= "TopBottomLeftRight")
#' 
#' ## Red, size 24, Bold, italic, underline, center aligned Font, bottom border
#' style <- createStyle(fontSize = 24, fontColour = rgb(1,0,0),
#'    textDecoration = c("bold", "italic", "underline"), 
#'    halign = "center", valign = "center", border = "Bottom")
#'  
#' # borderColour is recycled for each border or all colours can be supplied
#' 
#' # colour is recycled 3 times for "Top", "Bottom" & "Right" sides.
#' createStyle(border = "TopBottomRight", borderColour = "red") 
#' 
#' # supply all colours
#' createStyle(border = "TopBottomLeft", borderColour = c("red","yellow", "green"))
createStyle <- function(fontName = NULL,
                        fontSize = NULL,
                        fontColour = NULL,
                        numFmt = "GENERAL",
                        border = NULL,
                        borderColour = getOption("openxlsx.borderColour", "black"),
                        borderStyle =  getOption("openxlsx.borderStyle", "thin"),
                        bgFill = NULL, fgFill = NULL,
                        halign = NULL, valign = NULL, 
                        textDecoration = NULL, wrapText = FALSE,
                        textRotation = NULL){
  
  ### Error checking
  
  ## if num fmt is made up of dd, mm, yy
  
  numFmt <- tolower(numFmt[[1]])
  validNumFmt <- c("general", "number", "currency", "accounting", "date", "longdate", "time", "percentage", "scientific", "text", "3", "4", "comma")
  
  if(numFmt == "date"){
    numFmt <- getOption("openxlsx.dateFormat", getOption("openxlsx.dateformat", "date"))
  }else if(!numFmt %in% validNumFmt){
    if(grepl("[^mdyhsap[[:punct:] 0\\.#\\$\\*]", numFmt))
      stop("Invalid numFmt")
  }
  
  if(numFmt == "longdate"){
    numFmt <- getOption("openxlsx.datetimeFormat", getOption("openxlsx.datetimeformat", getOption("openxlsx.dateTimeFormat", "longdate")))  
  }
  
  
  numFmtMapping <- list(list("numFmtId" = 0),  # GENERAL
                        list("numFmtId" = 2),  # NUMBER
                        list("numFmtId" = 164, formatCode = "&quot;$&quot;#,##0.00"), ## CURRENCY
                        list("numFmtId" = 44), # ACCOUNTING
                        list("numFmtId" = 14), # DATE
                        list("numFmtId" = 166, formatCode = "yyyy/mm/dd hh:mm:ss"), #LONGDATE
                        list("numFmtId" = 167), # TIME
                        list("numFmtId" = 10),  # PERCENTAGE
                        list("numFmtId" = 11),  # SCIENTIFIC
                        list("numFmtId" = 49),
                        
                        list("numFmtId" = 3),
                        list("numFmtId" = 4),
                        list("numFmtId" = 3))
  
  names(numFmtMapping) <- validNumFmt
  
  ## Validate border line style
  if(!is.null(borderStyle))
    borderStyle <- validateBorderStyle(borderStyle)
  
  if(!is.null(halign)){
    halign <- tolower(halign[[1]])
    if(!halign %in% c("left", "right", "center"))
      stop("Invalid halign argument!")
  }
  
  if(!is.null(valign)){
    valign <- tolower(valign[[1]])
    if(!valign %in% c("top", "bottom", "center"))
      stop("Invalid valign argument!")
  }
  
  if(!is.logical(wrapText))
    stop("Invalid wrapText")
  
  textDecoration <- tolower(textDecoration)
  if(!is.null(textDecoration)){
    if(!all(textDecoration %in% c("bold", "strikeout", "italic", "underline", "underline2", "")))
      stop("Invalid textDecoration!")
  }
  
  borderColour <- validateColour(borderColour, "Invalid border colour!")
  
  if(!is.null(fontColour))
    fontColour <- validateColour(fontColour, "Invalid font colour!")
  
  if(!is.null(fontSize))
    if(fontSize < 1) stop("Font size must be greater than 0!")
  
  
  ######################### error checking complete #############################
  style <- Style$new()
  
  style$fontName <- list(val = fontName)
  style$fontSize <- list(val = fontSize)
  if(!is.null(fontColour))
    style$fontColour <- list("rgb" =  fontColour)
  
  style$fontDecoration <- toupper(textDecoration)
  
  ## background fill   
  if(is.null(bgFill)){
    bgFillList <- NULL
  }else{
    bgFill <- validateColour(bgFill, "Invalid bgFill colour")
    style$fill <- append(style$fill, list(fillBg = list("rgb" = bgFill)))
  }
  
  ## foreground fill
  if(is.null(fgFill)){
    fgFillList <- NULL
  }else{
    fgFill <- validateColour(fgFill, "Invalid fgFill colour")
    style$fill <- append(style$fill, list(fillFg = list(rgb = fgFill)))
  }
  
  
  ## border
  if(!is.null(border)){
    border <- toupper(border)
    
    ## find position of each side in string
    sides <- c("LEFT", "RIGHT", "TOP", "BOTTOM")
    pos <- sapply(sides, function(x) regexpr(x, border))
    pos <- pos[order(pos, decreasing = FALSE)]
    nSides <- sum(pos > 0)
    
    borderColour <- rep(borderColour, length.out = nSides)
    borderStyle <-  rep(borderStyle, length.out = nSides)
    
    pos <- pos[pos > 0]
    
    if(length(pos) == 0)
      stop("Unknown border argument")
    
    names(borderColour) <- names(pos)
    names(borderStyle) <- names(pos)
    
    if("LEFT" %in% names(pos)){
      style$borderLeft <- borderStyle[["LEFT"]]
      style$borderLeftColour <- list("rgb" = borderColour[["LEFT"]])
    }
    
    if("RIGHT" %in% names(pos)){
      style$borderRight <-  borderStyle[["RIGHT"]]
      style$borderRightColour <- list("rgb" = borderColour[["RIGHT"]])
    }
    
    if("TOP" %in% names(pos)){
      style$borderTop <-  borderStyle[["TOP"]]
      style$borderTopColour <- list("rgb" = borderColour[["TOP"]])
    }
    
    if("BOTTOM" %in% names(pos)){
      style$borderBottom <-  borderStyle[["BOTTOM"]]
      style$borderBottomColour <- list("rgb" = borderColour[["BOTTOM"]])
    }
    
  }
  
  ## other fields
  style$halign <- halign
  style$valign <- valign
  style$wrapText <- wrapText[[1]]
  
  if(!is.null(textRotation)){
    if(!is.numeric(textRotation))
      stop("textRotation must be numeric.")
    style$textRotation <- round(textRotation[[1]], 0)
  }
  
  if(numFmt %in% validNumFmt){
    style$numFmt <- numFmtMapping[[numFmt[[1]]]]
  }else{
    style$numFmt <- list("numFmtId" = 9999, formatCode = numFmt)  ## Custom numFmt
  }
  
  return(style)
} 



#' @name addStyle 
#' @title Add a style to a set of cells
#' @description Function adds a style to a specified set of cells.
#' @author Alexander Walker
#' @param wb A Workbook object containing a worksheet.
#' @param sheet A worksheet to apply the style to.
#' @param style A style object returned from createStyle()
#' @param rows Rows to apply style to.
#' @param cols columns to apply style to.
#' @param gridExpand If TRUE, style will be applied to all combinations of rows and cols.
#' @seealso \code{\link{createStyle}}
#' @seealso expand.grid
#' @export
#' @examples
#' ## See package vignette for more examples.
#' 
#' ## Create a new workbook
#' wb <- createWorkbook("My name here")
#' 
#' ## Add a worksheets
#' addWorksheet(wb, "Expenditure", gridLines = FALSE) 
#' 
#' ##write data to worksheet 1
#' writeData(wb, sheet = 1, USPersonalExpenditure, rowNames = TRUE)
#' 
#' ## create and add a style to the column headers
#' headerStyle <- createStyle(fontSize = 14, fontColour = "#FFFFFF", halign = "center",
#'                         fgFill = "#4F81BD", border="TopBottom", borderColour = "#4F81BD")
#' 
#' addStyle(wb, sheet = 1, headerStyle, rows = 1, cols = 1:6, gridExpand = TRUE)
#' 
#' ## style for body 
#' bodyStyle <- createStyle(border="TopBottom", borderColour = "#4F81BD")
#' addStyle(wb, sheet = 1, bodyStyle, rows = 2:6, cols = 1:6, gridExpand = TRUE)
#' setColWidths(wb, 1, cols=1, widths = 21) ## set column width for row names column
#' 
#' saveWorkbook(wb, "addStyleExample.xlsx", overwrite = TRUE)
addStyle <- function(wb, sheet, style, rows, cols, gridExpand = FALSE){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!"Style" %in% class(style))
    stop("style argument must be a Style object.")
  
  cols <- convertFromExcelRef(cols)
  rows <- as.integer(rows)
  
  ## rows and cols need to be the same length
  if(gridExpand){
    combs <- expand.grid(cols, rows) 
    cols <- combs[,1]
    rows <- combs[,2]
  }
  
  if(length(rows) != length(cols)){
    stop("Length of rows and cols must be equal.")
  }
  
  styleElements <- list(style = style,
                        cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                          rows = rows,
                                          cols = cols)))
  
  invisible(wb$styleObjects <- append(wb$styleObjects, list(styleElements)))
  
}


#' @name getCellRefs
#' @title Return excel cell coordinates from (x,y) coordinates
#' @author Alexander Walker
#' @param cellCoords A data.frame with two columns coordinate pairs. 
#' @return Excel alphanumeric cell reference
getCellRefs <- function(cellCoords){
  l <- .Call("openxlsx_convert2ExcelRef", unlist(cellCoords[,2]), LETTERS, PACKAGE="openxlsx")
  paste0(l, cellCoords[,1])
}


#' @name conditionalFormat
#' @title Add conditional formatting to cells
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param rule The condition under which to apply the formatting or a vector of colours. See examples.
#' @param style A style to apply to those cells that satisify the rule. A Style object returned from createStyle()
#' @details Valid operators are "<", "<=", ">", ">=", "==", "!=". See Examples.
#' Default style given by: createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' @param type Either 'expression', 'colorscale' or 'databar'. If 'expression' the formatting is determined
#' by a formula.  If colorScale cells are coloured based on cell value. See examples.
#' @seealso \code{\link{createStyle}}
#' @export
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "cellIs")
#' addWorksheet(wb, "moving Row")
#' addWorksheet(wb, "moving Col")
#' addWorksheet(wb, "Dependent on 1")
#' addWorksheet(wb, "colourScale 2 Colours")
#' addWorksheet(wb, "databar")
#' 
#' negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
#' 
#' ## rule applies to all each cell in range
#' writeData(wb, 1, -5:5)
#' writeData(wb, 1, LETTERS[1:11], startCol=2)
#' conditionalFormat(wb, 1, cols=1, rows=1:11, rule="!=0", style = negStyle)
#' conditionalFormat(wb, 1, cols=1, rows=1:11, rule="==0", style = posStyle)
#' 
#' ## highlight row dependent on first cell in row
#' writeData(wb, 2, -5:5)
#' writeData(wb, 2, LETTERS[1:11], startCol=2)
#' conditionalFormat(wb, 2, cols=1:2, rows=1:11, rule="$A1<0", style = negStyle)
#' conditionalFormat(wb, 2, cols=1:2, rows=1:11, rule="$A1>0", style = posStyle)
#' 
#' ## highlight column dependent on first cell in column
#' writeData(wb, 3, -5:5)
#' writeData(wb, 3, LETTERS[1:11], startCol=2)
#' conditionalFormat(wb, 3, cols=1:2, rows=1:11, rule="A$1<0", style = negStyle)
#' conditionalFormat(wb, 3, cols=1:2, rows=1:11, rule="A$1>0", style = posStyle)
#' 
#' 
#' ## highlight entire range cols X rows dependent only on cell A1
#' writeData(wb, 4, -5:5)
#' writeData(wb, 4, LETTERS[1:11], startCol=2)
#' conditionalFormat(wb, 4, cols=1:2, rows=1:11, rule="$A$1<0", style = negStyle)
#' conditionalFormat(wb, 4, cols=1:2, rows=1:11, rule="$A$1>0", style = posStyle)
#' 
#' ## colourscale colours cells based on cell value
#' 
#' df <- read.xlsx(system.file("readTest.xlsx", package = "openxlsx"), sheet = 5)
#' writeData(wb, 5, df, colNames=FALSE)  ## write data.frame
#' 
#' ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
#' conditionalFormat(wb, 5, cols=1:ncol(df), rows=1:nrow(df),
#'    rule =c("black", "white"), type = "colourScale")
#'    
#' setColWidths(wb, 5, cols=1:ncol(df), widths=1.07)
#' setRowHeights(wb, 5, rows=1:nrow(df), heights=7.5) 
#'
#' ## Databars
#' writeData(wb, "databar", -5:5)
#' conditionalFormat(wb, "databar", cols = 1, rows = 1:12, type = "databar") ## Default colours
#' 
#' writeData(wb, "databar", -5:5, startCol = 2)
#' ## set negative and positive colours
#' conditionalFormat(wb, "databar", cols = 2, rows = 1:12,
#'  rule = c("yellow", "green"), type = "databar")
#' 
#' ## Save workbook
#' saveWorkbook(wb, "conditionalFormatExample.xlsx", overwrite = TRUE)
conditionalFormat <- function(wb, sheet, cols, rows, rule = NULL, style = NULL, type = "expression"){
  
  
  ## Rule always applies to top left of sqref, $ determine which cells the rule depends on
  ## Rule for "databar" and colourscale are colours of length 2/3 or 1 respectively.
  
  type <- tolower(type)
  if(tolower(type) %in% c("colorscale", "colourscale")){
    type <- "colorScale"
  }else if(type == "databar"){
    type <- "dataBar"
  }else if(type != "expression"){
    stop("Invalid type argument.  Type must be 'expression', 'colourScale' or 'databar'")
  }
  
  ## rows and cols
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)  
  rows <- as.integer(rows)
  
  ## check valid rule
  if(type == "colorScale"){
    if(!length(rule) %in% 2:3)
      stop("rule must be a vector containing 2 or 3 colours if type is 'colorScale'")
    
    rule <- validateColour(rule, errorMsg="Invalid colour specified in rule.")
    dxfId <- NULL
    
  }else if(type == "dataBar"){
    
    ## If rule is NULL use default colour
    if(is.null(rule)){
      rule <- "FF638EC6"
    }else{
      rule <- validateColour(rule, errorMsg="Invalid colour specified in rule.")
    }
    
    dxfId <- NULL
    
  }else{ ## else type == "expression"
    
    rule <- toupper(gsub(" ", "", rule))
    rule <- replaceIllegalCharacters(rule)
    rule <- gsub("!=", "&lt;&gt;", rule)
    rule <- gsub("==", "=", rule)
    
    if(!grepl("[A-Z]", substr(rule, 1, 2))){
      
      ## formula looks like "operatorX" , attach top left cell to rule    
      rule <- paste0( getCellRefs(data.frame("x" = min(rows), "y" = min(cols))), rule)
      
    } ## else, there is a letter in the formula and apply as is
    
    if(is.null(style))
      style <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
    
    invisible(dxfId <- wb$addDXFS(style))
    
  }
  
  
  invisible(wb$conditionalFormatCell(sheet,
                                     startRow = min(rows),
                                     endRow = max(rows),
                                     startCol = min(cols),
                                     endCol = max(cols),
                                     dxfId,
                                     formula = rule,
                                     type = type))
  
  invisible(0)
  
}



#' @name freezePane
#' @title Freeze a worksheet pane
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param firstActiveRow Top row of active region
#' @param firstActiveCol Furthest left column of active region
#' @param firstRow If TRUE, freezes the first row (equivalent to firstActiveRow = 2)
#' @param firstCol If TRUE, freezes the first column (equivalent to firstActiveCol = 2)
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook("Kenshin")
#' 
#' ## Add some worksheets
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#' addWorksheet(wb, "Sheet 3")
#' addWorksheet(wb, "Sheet 4")
#'
#' ## Freeze Panes
#' freezePane(wb, "Sheet 1" ,  firstActiveRow = 5,  firstActiveCol = 3)
#' freezePane(wb, "Sheet 2", firstCol = TRUE)  ## shortcut to firstActiveCol = 2
#' freezePane(wb, 3, firstRow = TRUE)  ## shortcut to firstActiveRow = 2
#' freezePane(wb, 4, firstActiveRow = 1, firstActiveCol = "D")
#'  
#' ## Save workbook
#' saveWorkbook(wb, "freezePaneExample.xlsx", overwrite = TRUE)
freezePane <- function(wb, sheet, firstActiveRow = NULL, firstActiveCol = NULL, firstRow = FALSE, firstCol = FALSE){
  
  
  ## Convert to numeric if column letter given
  firstActiveRow <- convertFromExcelRef(firstActiveRow)
  firstActiveCol <- convertFromExcelRef(firstActiveCol)
  
  if(is.null(firstActiveRow)) firstActiveRow <- 1L
  if(is.null(firstActiveCol)) firstActiveCol <- 1L
  if(!is.logical(firstRow)) firstRow <- FALSE
  if(!is.logical(firstCol)) firstCol <- FALSE
  
  if(firstRow & !firstCol){
    invisible(wb$freezePanes(sheet, firstRow = firstRow))
  }else if(firstCol &! firstRow){
    invisible(wb$freezePanes(sheet, firstCol = firstCol))
  }else if(firstRow & firstCol){
    invisible(wb$freezePanes(sheet, firstActiveRow = 2L, firstActiveCol = 2L))
  }else{ 
    
    if(!is.numeric(firstActiveRow))
      stop("firstActiveRow must be numeric.")
    
    invisible(wb$freezePanes(sheet, firstActiveRow = firstActiveRow, firstActiveCol = firstActiveCol, firstRow = firstRow, firstCol = firstCol)  )
  }
  
}


convert2EMU <- function(d, units){
  
  if(grepl("in", units))
    d <- d*2.54
  
  if(grepl("mm|milli", units))
    d <- d/10
  
  return(d*360000)
  
}




#' @name insertImage
#' @title Insert an image into a worksheet 
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param file An image file. Valid file types are: jpeg, png, bhp
#' @param width Width of figure.
#' @param height Height of figure.
#' @param startRow Row coordinate of upper left corner of the image
#' @param startCol Column coordinate of upper left corner of the image
#' @param units Units of width and height. Can be "in", "cm" or "px"
#' @param dpi Image resolution used for conversion between units.
#' @seealso \code{\link{insertPlot}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook("Ayanami")
#' 
#' ## Add some worksheets
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#' addWorksheet(wb, "Sheet 3")
#'
#' ## Insert images
#' img <- system.file("einstein.jpg", package = "openxlsx")
#' insertImage(wb, "Sheet 1", img, startRow = 5,  startCol = 3, width = 6, height = 5)
#' insertImage(wb, 2, img, startRow = 2,  startCol = 2)
#' insertImage(wb, 3 , img, width = 15, height = 12, startRow = 3, startCol = "G", units = "cm")
#'  
#' ## Save workbook
#' saveWorkbook(wb, "insertImageExample.xlsx", overwrite = TRUE)
insertImage <- function(wb, sheet, file, width = 6, height = 3, startRow = 1, startCol = 1, units = "in", dpi = 300){
  
  if(!file.exists(file))
    stop("File does not exist.")
  
  if(!grepl("\\\\|\\/", file))
    file <- file.path(getwd(), file, fsep = .Platform$file.sep)
  
  units <- tolower(units)
  
  if(!units %in% c("cm", "in", "px"))
    stop("Invalid units.\nunits must be one of: cm, in, px")
  
  startCol <- convertFromExcelRef(startCol)
  startRow <- as.integer(startRow)
  
  ##convert to inches
  if(units == "px"){
    width <- width/dpi
    height <- height/dpi
  }else if(units == "cm"){
    width <- width/2.54
    height <- height/2.54
  }
  
  ## Convert to EMUs
  widthEMU <- as.integer(round(width * 914400L, 0)) #(EMUs per inch)
  heightEMU <- as.integer(round(height * 914400L, 0)) #(EMUs per inch)
  
  wb$insertImage(sheet, file = file, startRow = startRow, startCol = startCol, width = widthEMU, height = heightEMU)
  
}

pixels2ExcelColWidth <- function(pixels){
  
  if(any(!is.numeric(pixels)))
    stop("All elements of pixels must be numeric")
  
  pixels[pixels == 0] <- 8.43
  pixels[pixels != 0] <- (pixels[pixels != 0] - 12) / 7 +  1
  
  pixels
}


#' @name setRowHeights
#' @title Set worksheet row heights
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Indices of rows to set height
#' @param heights Heights to set rows to specified in Excel column height units.
#' @seealso \code{\link{removeRowHeights}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Add a worksheet
#' addWorksheet(wb, "Sheet 1") 
#'
#' ## set row heights
#' setRowHeights(wb, 1, rows = c(1,4,22,2,19), heights = c(24,28,32,42,33))
#' 
#' ## overwrite row 1 height
#' setRowHeights(wb, 1, rows = 1, heights = 40)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "setRowHeightsExample.xlsx", overwrite = TRUE)
setRowHeights <- function(wb, sheet, rows, heights){
  
  sheet <- wb$validateSheet(sheet)
  
  if(length(rows) > length(heights))
    heights <- rep(heights, length.out = length(rows))
  
  if(length(heights) > length(rows))
    stop("Greater number of height values than rows.")
  
  ## Remove duplicates
  heights <- heights[!duplicated(rows)]
  rows <- rows[!duplicated(rows)]
  
  
  heights <- as.list(as.character(as.numeric(heights)))
  names(heights) <- rows
  
  wb$setRowHeights(sheet, rows, heights)
  
}

#' @name setColWidths
#' @title Set worksheet column widths
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of cols to set width
#' @param widths widths to set rows to specified in Excel column width units.
#' @seealso \code{\link{removeColWidths}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Add a worksheet
#' addWorksheet(wb, "Sheet 1") 
#'
#' ## set col widths
#' setColWidths(wb, 1, cols = c(1,4,6,7,9), widths = c(16,15,12,18,33))
#' 
#' ## Save workbook
#' saveWorkbook(wb, "setColWidthsExample.xlsx", overwrite = TRUE)
setColWidths <- function(wb, sheet, cols, widths){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  widths <- tolower(widths)  ## possibly "auto"
  
  if(length(widths) > length(cols))
    stop("More widths than columns supplied.")
  
  if(length(widths) < length(cols))
    widths <- rep(widths, length.out = length(cols))
  
  ## check for existing custom widths
  exColsWidths <- unlist(lapply(wb$colWidths[[sheet]], "[[", "col"))
  flag <- exColsWidths %in% cols
  if(any(flag))
    wb$colWidths[[sheet]] <- wb$colWidths[[sheet]][!flag]
  
  cols <- convertFromExcelRef(cols)
  
  ## Remove duplicates
  widths <- widths[!duplicated(cols)]
  cols <- cols[!duplicated(cols)]
  
  allWidths <- append(wb$colWidths[[sheet]], lapply(1:length(cols), function(i) c('col' = cols[[i]], 'width' = widths[[i]])))
  allWidths <- allWidths[order(as.integer(sapply(allWidths, "[[", "col")))]
  
  wb$colWidths[[sheet]] <- allWidths
}


#' @name removeColWidths
#' @title Remove custom column widths from a worksheet
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of colunss to remove custom width (if any) from.
#' @seealso \code{\link{setColWidths}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))
#' 
#' ## remove column widths in columns 1 to 10
#' removeColWidths(wb, 1, cols = 1:10)
#' saveWorkbook(wb, "removeColWidthsExample.xlsx", overwrite = TRUE)
removeColWidths <- function(wb, sheet, cols){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)
  
  customCols <- as.integer(unlist(lapply(wb$colWidths[[sheet]], "[[", "col")))
  removeInds <- which(customCols %in% cols)
  if(length(removeInds) > 0)
    wb$colWidths[[sheet]] <- wb$colWidths[[sheet]][-removeInds]
  
}



#' @name removeRowHeights
#' @title Remove custom row heights from a worksheet
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Indices of rows to remove custom height (if any) from.
#' @seealso \code{\link{setRowHeights}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))
#'
#' ## remove any custom row heights in rows 1 to 10
#' removeRowHeights(wb, 1, rows = 1:10)
#' saveWorkbook(wb, "removeRowHeightsExample.xlsx", overwrite = TRUE)
removeRowHeights <- function(wb, sheet, rows){
  
  sheet <- wb$validateSheet(sheet)
  
  customRows <- as.integer(names(wb$rowHeights[[sheet]]))
  removeInds <- which(customRows %in% rows)
  if(length(removeInds) > 0)
    wb$rowHeights[[sheet]] <- wb$rowHeights[[sheet]][-removeInds]
  
}


#' @name insertPlot
#' @title Insert the current plot into a worksheet
#' @author Alexander Walker
#' @description The current plot is saved to a temporary image file using dev.copy.  
#' This file is then written to the workbook using insertImage.
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param startRow Row coordinate of upper left corner of figure. xy[[2]] when xy is given.
#' @param startCol Column coordinate of upper left corner of figure. xy[[1]] when xy is given.
#' @param xy Alternate way to specify startRow and startCol.  A vector of length 2 of form (startcol, startRow)
#' @param width Width of figure. Defaults to 6in.
#' @param height Height of figure . Defaults to 4in.
#' @param fileType File type of image
#' @param units Units of width and height. Can be "in", "cm" or "px"
#' @param dpi Image resolution
#' @seealso \code{\link{insertImage}}
#' @export
#' @examples
#' \dontrun{
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Add a worksheet
#' addWorksheet(wb, "Sheet 1", gridLines = FALSE) 
#'
#' ## create plot objects
#' require(ggplot2)
#' p1 <- qplot(mpg, data=mtcars, geom="density",
#'   fill=as.factor(gear), alpha=I(.5), main="Distribution of Gas Mileage")
#' p2 <- qplot(age, circumference,
#'   data = Orange, geom = c("point", "line"), colour = Tree)
#' 
#' ## Insert currently displayed plot to sheet 1, row 1, column 1
#' print(p1) #plot needs to be showing
#' insertPlot(wb, 1, width = 5, height = 3.5, fileType = "png", units = "in")
#' 
#' ## Insert plot 2
#' print(p2)
#' insertPlot(wb, 1, xy = c("J", 2), width = 16, height = 10,  fileType = "png", units = "cm")
#'
#' ## Save workbook
#' saveWorkbook(wb, "insertPlotExample.xlsx", overwrite = TRUE)
#' }
insertPlot <- function(wb, sheet, width = 6, height = 4, xy = NULL,
                       startRow = 1, startCol = 1, fileType = "png", units = "in", dpi = 300){
  
  
  if(is.null(dev.list()[[1]])){
    warning("No plot to insert.")
    return()
  }
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!is.null(xy)){
    startCol <- xy[[1]]
    startRow <- xy[[2]]
  }
  
  fileType <- tolower(fileType)
  units <- tolower(units)
  
  if(fileType == "jpg")
    fileType = "jpeg"
  
  if(!fileType %in% c("png", "jpeg", "tiff", "bmp"))
    stop("Invalid file type.\nfileType must be one of: png, jpeg, tiff, bmp")
  
  if(!units %in% c("cm", "in", "px"))
    stop("Invalid units.\nunits must be one of: cm, in, px")
  
  fileName <- tempfile(pattern = "figureImage", fileext= paste0(".", fileType))
  
  if(fileType == "bmp"){
    dev.copy(bmp, filename = fileName, width = width, height = height, units = units, res = dpi)
  }else if(fileType == "jpeg"){
    dev.copy(jpeg, filename = fileName, width = width, height = height, units = units, quality = 100, res = dpi)
  }else if(fileType == "png"){
    dev.copy(png, filename = fileName, width = width, height = height, units = units, res = dpi)
  }else if(fileType == "tiff"){
    dev.copy(tiff, filename = fileName, width = width, height = height, units = units, compression = "none")
  }
  
  ## write image
  invisible(dev.off())
  
  insertImage(wb = wb, sheet = sheet, file = fileName, width = width, height = height, startRow = startRow, startCol = startCol, units = units, dpi = dpi)
  
}



#' @name replaceStyle
#' @title Replace an existing cell style
#' @author Alexander Walker
#' @param wb A workbook object
#' @param index Index of style object to replace
#' @param newStyle A style to replace the exising style as position index
#' @description Replace a style object
#' @export
#' @seealso \code{\link{getStyles}}
#' @examples
#' ## load a workbook 
#' wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))
#' 
#' ## create a new style and replace style 2
#' 
#' newStyle <- createStyle(fgFill = "#00FF00")
#'  
#' ## replace style 2
#' getStyles(wb) ## prints styles
#' replaceStyle(wb, 2, newStyle = newStyle)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "replaceStyleExample.xlsx", overwrite = TRUE)
replaceStyle <- function(wb, index, newStyle){
  
  nStyles <- length(wb$styleObjects)
  
  if(nStyles == 0)
    stop("Workbook has no existing styles.")
  
  if(index > nStyles)
    stop(sprintf("Invalid index. Workbook only has %s styles.", nStyles))
  
  if(!all("Style" %in% class(newStyle)))
    stop("Invalid style object.")
  
  wb$styleObjects[[index]]$style <- newStyle
  
}


#' @name getStyles
#' @title Returns a list of all styles in the workbook
#' @author Alexander Walker
#' @param wb A workbook object
#' @export
#' @seealso \code{\link{replaceStyle}}
#' @examples
#' ## load a workbook 
#' wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))
#' getStyles(wb)
getStyles <- function(wb){
  
  nStyles <- length(wb$styleObjects)
  
  if(nStyles == 0)
    stop("Workbook has no existing styles.")
  
  styles <- lapply(wb$styleObjects, "[[", "style")
  
  return(styles)
}



#' @name removeWorksheet
#' @title Remove a worksheet from a workbook
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @description Remove a worksheet from a workbook
#' @export
#' @examples
#' ## load a workbook 
#' wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))
#' 
#' ## Remove sheet 2
#' removeWorksheet(wb, 2)
#' 
#' ## save the modified workbook
#' saveWorkbook(wb, "removeWorksheetExample.xlsx", overwrite = TRUE)
removeWorksheet <- function(wb, sheet){
  
  if(class(wb) != "Workbook")
    stop("wb must be a Workbook object!")
  
  wb$deleteWorksheet(sheet)
  invisible(NULL)
}


#' @name deleteData
#' @title Delete cell data
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Rows to delete data from.
#' @param cols columns to delete data from.
#' @param gridExpand If TRUE, all data in rectangle min(rows):max(rows) X min(cols):max(cols)
#' will be removed.
#' @description Remove contents and styling from a cell.
#' @export
#' @examples
#' ## load a workbook 
#' wb <- createWorkbook()
#' addWorksheet(wb, "Worksheet 1")
#' x <- data.frame(matrix(runif(200), ncol = 10)) 
#' names(x) <- paste("Variable", 1:10)
#' writeData(wb, sheet = 1, x = x, startCol = 2, startRow = 3, colNames = TRUE)
#' 
#' ## delete cell contents
#' deleteData(wb, sheet = 1, cols = 3:5, rows = 5:7, gridExpand = TRUE)
#' deleteData(wb, sheet = 1, cols = 7:9, rows = 5:7, gridExpand = TRUE)
#' 
#' saveWorkbook(wb, "deleteDataExample.xlsx", overwrite = TRUE)
deleteData <- function(wb, sheet, cols, rows, gridExpand = FALSE){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  cols <- convertFromExcelRef(cols)
  rows <- as.integer(rows)
  
  ## rows and cols need to be the same length
  if(gridExpand){
    combs <- expand.grid(rows, cols) 
    rows <- combs[,1]
    cols <- combs[,2]
  }
  
  if(length(rows) != length(cols)){
    stop("Length of rows and cols must be equal.")
  }
  
  comb <- paste0(.Call('openxlsx_convert2ExcelRef', cols, LETTERS, PACKAGE="openxlsx"), rows)
  
  cellRefs <- sapply(wb$sheetData[[sheet]], "[[", "r")  
  wb$sheetData[[sheet]] <- wb$sheetData[[sheet]][!cellRefs %in% comb]
  
  invisible(1)
}


#' @name modifyBaseFont
#' @title Modify the default font
#' @author Alexander Walker
#' @param wb A workbook object
#' @param fontSize font size
#' @param fontColour font colour
#' @param fontName Name of a font
#' @description Base font is black, size 11, Calibri
#' @details  The font name is not validated in anyway.  Excel replaces unknown font names
#' with Arial.
#' @export
#' @examples
#' ## create a workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' ## modify base font to size 10 Arial Narrow in red
#' modifyBaseFont(wb, fontSize = 10, fontColour = "#FF0000", fontName = "Arial Narrow")
#' 
#' writeData(wb, "S1", iris)
#' writeDataTable(wb, "S1", x = iris, startCol = 10) ## font colour does not affect tables
#' saveWorkbook(wb, "modifyBaseFontExample.xlsx", overwrite = TRUE)
modifyBaseFont <- function(wb, fontSize = 11, fontColour = "#000000", fontName = "Calibri"){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  
  if(fontSize < 0) stop("Invalid fontSize")
  fontColour <- validateColour(fontColour)
  
  wb$styles$fonts[[1]] <- sprintf('<font><sz val="%s"/><color rgb="%s"/><name val="%s"/></font>', fontSize, fontColour, fontName)
  
}


#' @name getBaseFont
#' @title Return the workbook defaul font
#' @author Alexander Walker
#' @param wb A workbook object
#' @description Returns the base font used in the workbook.
#' @export
#' @examples
#' ## create a workbook
#' wb <- createWorkbook()
#' getBaseFont(wb)
#' 
#' ## modify base font to size 10 Arial Narrow in red
#' modifyBaseFont(wb, fontSize = 10, fontColour = "#FF0000", fontName = "Arial Narrow")
#' 
#' getBaseFont(wb)
getBaseFont <- function(wb){
  
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  wb$getBaseFont()
  
}


#' @name setHeader
#' @title Set header for all worksheets
#' @author Alexander Walker
#' @param wb A workbook object
#' @param text header text. A character vector of length 1.
#' @param position Postion of text in header. One of "left", "center" or "right"
#' @export
#' @examples
#' wb <- createWorkbook("Edgar Anderson")
#' addWorksheet(wb, "S1")
#' writeDataTable(wb, "S1", x = iris[1:30,], xy = c("C", 5))
#' 
#' ## set all headers
#' setHeader(wb, "This is a header", position="center")
#' setHeader(wb, "To the left", position="left")
#' setHeader(wb, "On the right", position="right")
#' 
#' ## set all footers
#' setFooter(wb, "Center Footer Here", position="center")
#' setFooter(wb, "Bottom left", position="left")
#' setFooter(wb, Sys.Date(), position="right")
#' 
#' saveWorkbook(wb, "headerFooterExample.xlsx", overwrite = TRUE)
setHeader <- function(wb, text, position = "center"){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  position <- tolower(position)
  if(!position %in% c("left", "center", "right")) 
    stop("Invalid position.")
  
  if(length(text) != 1)
    stop("Text argument must be a character vector of length 1")
  
  sheet <- wb$validateSheet(1)
  wb$headFoot$text[wb$headFoot$pos == position & wb$headFoot$head == "head"] <- as.character(text)
  
}


#' @name setFooter
#' @title Set footer for all worksheets
#' @author Alexander Walker
#' @param wb A workbook object
#' @param text footer text. A character vector of length 1.
#' @param position Postion of text in footer. One of "left", "center" or "right"
#' @export
#' @examples
#' wb <- createWorkbook("Edgar Anderson")
#' addWorksheet(wb, "S1")
#' writeDataTable(wb, "S1", x = iris[1:30,], xy = c("C", 5))
#' 
#' ## set all headers
#' setHeader(wb, "This is a header", position="center")
#' setHeader(wb, "To the left", position="left")
#' setHeader(wb, "On the right", position="right")
#' 
#' ## set all footers
#' setFooter(wb, "Center Footer Here", position="center")
#' setFooter(wb, "Bottom left", position="left")
#' setFooter(wb, Sys.Date(), position="right")
#' 
#' saveWorkbook(wb, "headerFooterExample.xlsx", overwrite = TRUE)
setFooter <- function(wb, text, position = "center"){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  position <- tolower(position)
  if(!position %in% c("left", "center", "right")) 
    stop("Invalid position.")
  
  if(length(text) != 1)
    stop("Text argument must be a character vector of length 1")
  
  sheet <- wb$validateSheet(1)
  wb$headFoot$text[wb$headFoot$pos == position & wb$headFoot$head == "foot"] <- as.character(text)
  
}





#' @name pageSetup
#' @title Set page margins, orientation and print scaling
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param orientation Page orientation. One of "portrait" or "landscape"
#' @param scale Print scaling. Numeric value between 10 and 400
#' @param left left page margin in inches
#' @param right right page margin in inches
#' @param top top page margin in inches
#' @param bottom bottom page margin in inches
#' @param header header margin in inches
#' @param footer footer margin in inches
#' @export
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' addWorksheet(wb, "S2")
#' writeDataTable(wb, 1, x = iris[1:30,])
#' writeDataTable(wb, 2, x = iris[1:30,], xy = c("C", 5))
#' 
#' ## landscape page scaled to 50%
#' pageSetup(wb, sheet = 1, orientation = "landscape", scale = 50)
#' 
#' ## portrait page scales to 300% with 0.5in left and right margins
#' pageSetup(wb, sheet = 2, orientation = "portrait", scale = 300, left= 0.5, right = 0.5)
#' 
#' saveWorkbook(wb, "pageSetupExample.xlsx", overwrite = TRUE)
pageSetup <- function(wb, sheet, orientation = "portrait", scale = 100, left = 0.7, right = 0.7, top = 0.75, bottom = 0.75, header = 0.3, footer = 0.3){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  orientation <- tolower(orientation)
  if(!orientation %in% c("portrait", "landscape")) stop("Invalid page orientation.")
  
  if(scale < 10 | scale > 400)
    stop("Scale must be between 10 and 400.")
  
  sheet <- wb$validateSheet(sheet)
  
  wb$worksheets[[sheet]]$pageSetup <- sprintf('<pageSetup paperSize="9" orientation="%s" scale = "%s" horizontalDpi="300" verticalDpi="300" r:id="rId2"/>', 
                                              orientation, scale)
  
  wb$worksheets[[sheet]]$pageMargins <- 
    sprintf('<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>"', left, right, top, bottom, header, footer)
  
}




#' @name showGridLines
#' @title Set worksheet gridlines to show or hide.
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param showGridLines A logical. If TRUE, grid lines are hidden.
#' @export
#' @examples
#' wb <- loadWorkbook(xlsxFile = file.path(path.package("openxlsx"), "loadExample.xlsx"))
#' names(wb) ## list worksheets in workbook
#' showGridLines(wb, 1, showGridLines = FALSE)
#' showGridLines(wb, "Empty sheet", showGridLines = FALSE)
#' saveWorkbook(wb, "showGridLinesExample.xlsx", overwrite = TRUE)
showGridLines <- function(wb, sheet, showGridLines = FALSE){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  sheet <- wb$validateSheet(sheet)
  
  if(!is.logical(showGridLines)) stop("showGridLines must be a logical")
  
  
  sv <- wb$worksheets[[sheet]]$sheetViews
  showGridLines <- as.integer(showGridLines)
  ## If attribute exists gsub
  if(grepl("showGridLines", sv)){
    sv <- gsub('showGridLines=".?[^"]', sprintf('showGridLines="%s', showGridLines), sv, perl = TRUE)
  }else{
    sv <- gsub('<sheetView ', sprintf('<sheetView showGridLines="%s" ', showGridLines), sv)
  }
  
  wb$worksheets[[sheet]]$sheetViews <- sv
  
}





#' @name worksheetOrder
#' @title Order of worksheets in xlsx file
#' @details This function does not reorder the worksheets within the workbook object, it simply
#' shuffles the order when writing to file.
#' @export
#' @examples
#' ## setup a workbook with 3 worksheets
#' wb <- createWorkbook()
#' addWorksheet(wb = wb, sheetName = "Sheet 1", gridLines = FALSE)
#' writeDataTable(wb = wb, sheet = 1, x = iris)
#' 
#' addWorksheet(wb = wb, sheetName = "mtcars (Sheet 2)", gridLines = FALSE)
#' writeData(wb = wb, sheet = 2, x = mtcars)
#' 
#' addWorksheet(wb = wb, sheetName = "Sheet 3", gridLines = FALSE)
#' writeData(wb = wb, sheet = 3, x = Formaldehyde)
#' 
#' worksheetOrder(wb)
#' names(wb)
#' worksheetOrder(wb) <- c(1,3,2) # switch position of sheets 2 & 3 
#' writeData(wb, 2, 'This is still the "mtcars" worksheet', startCol = 15)
#' worksheetOrder(wb)
#' names(wb)  ## ordering within workbook is not changed
#' 
#' saveWorkbook(wb, "worksheetOrderExample.xlsx",  overwrite = TRUE)
#' worksheetOrder(wb) <- c(3,2,1)
#' saveWorkbook(wb, "worksheetOrderExample2.xlsx",  overwrite = TRUE)
worksheetOrder <- function(wb){
  
  if(!"Workbook" %in% class(wb))
    stop("Argument must be a Workbook.")
  
  #   nms <- names(wb$worksheets)
  #   nms <- replaceXMLEntities(nms)
  #   sprintf('%s: "%s"', wb$sheetOrder, nms[wb$sheetOrder])
  
  wb$sheetOrder
  
}

#' @rdname worksheetOrder
#' @param wb A workbook object
#' @param value Vector specifying order to write worksheets to file
#' @export
`worksheetOrder<-` <- function(wb, value) {
  
  if(!"Workbook" %in% class(wb))
    stop("Argument must be a Workbook.")
  
  value <- unique(value)
  if(length(value) != length(wb$worksheets))
    stop(sprintf("Worksheet order must be same length as number of worksheets [%s]", length(wb$worksheets)))
  
  if(any(value > length(wb$worksheets)))
    stop("Elements of order are greater than the number of worksheets")
  
  wb$sheetOrder <- value
  
  invisible(wb)
  
}




#' @name convertToDate
#' @title Convert from excel date number to R Date type
#' @param x A vector of integers
#' @param origin date. Default value is for Windows Excel 2010
#' @details Excel stores dates as number of days from some origin day
#' (this origin is "1970-1-1" for Excel 2010).
#' @export
#' @examples
#' ##2014 April 21st to 25th
#' x <- c(41750, 41751, 41752, 41753, 41754) 
#' convertToDate(x)
#' convertToDate(c(41821.8127314815, 41820.8127314815))
convertToDate <- function(x, origin = "1970-1-1"){
  as.Date(x - 25569, origin = origin)  
}


#' @name convertToDateTime
#' @title Convert from excel time number to R as.POSIXct
#' @param x A numeric vector
#' @param origin date. Default value is for Windows Excel 2010
#' @details Excel stores dates as number of days from some origin day
#' (this origin is "1970-1-1" for Excel 2010).
#' @export
#' @examples
#' ##2014 April 21st to 25th
#' x <- c(41821.8127314815, 41820.8127314815) 
#' convertToDateTime(x)
convertToDateTime <- function(x, origin = "1970-1-1"){
  
  rem <- x %% 1
  date <- as.Date(as.integer(x) - 25569, origin = origin)  
  fraction <- 24*rem
  hrs <- floor(fraction)
  minFrac <- (fraction-hrs)*60
  mins <- floor(minFrac)
  secs <- (minFrac - mins)*60
  y <- paste(hrs, mins, secs, sep = ":")
  y <- format(strptime(y, "%H:%M:%S"), "%H:%M:%S") 
  dateTime <- as.POSIXct(paste(date, y))
  
  return(dateTime)
}



#' @name names
#' @aliases names.Workbook
#' @export names.Workbook
#' @method names Workbook
#' @title get or set worksheet names
#' @param x A \code{Workbook} object
#' @examples
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' addWorksheet(wb, "S2")
#' addWorksheet(wb, "S3")
#' 
#' names(wb)
#' names(wb)[[2]] <- "S2a"
#' names(wb)
#' names(wb) <- paste("Sheet", 1:3)
names.Workbook <- function(x){
  nms <- names(x$worksheets)
  nms <- replaceXMLEntities(nms)
}

#' @rdname names
#' @param value a character vector the same length as wb
#' @export
`names<-.Workbook` <- function(x, value) {
  
  if(any(duplicated(value)))
    stop("Worksheet names must be unique.")
  
  exSheets <- names(x$worksheets)
  inds <- which(value != exSheets)
  
  if(length(inds) == 0)
    return(invisible(x))
  
  if(length(value) != length(x$worksheets))
    stop(sprintf("names vector must have length equal to number of worksheets in Workbook [%s]", length(exSheets)))
  
  if(any(nchar(value) > 31)){
    warning("Worksheet names must less than 32 characters. Truncating names...")
    value[nchar(value) > 31] <- sapply(value[nchar(value) > 31], substr, start = 1, stop = 31)
  }
  
  for(i in inds)
    invisible(x$setSheetName(i, value[[i]]))
  
  invisible(x)
  
}





#' @name addFilter
#' @title add filters to columns
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols columns to add filter to. 
#' @param rows A row numbers
#' @seealso \code{\link{writeData}}
#' @details adds filters to worksheet columns, same as filter parameters in writeData.
#' writeDataTable automatically adds filters to first row of a table.
#' NOTE Can only have a single filter per worksheet unless using tables.
#' @export
#' @seealso \code{\link{addFilter}}
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#' addWorksheet(wb, "Sheet 3")
#' 
#' writeData(wb, 1, iris)
#' addFilter(wb, 1, row = 1, cols = 1:ncol(iris))
#' 
#' ## Equivalently
#' writeData(wb, 2, x = iris, filter = TRUE)
#' 
#' ## Similarly
#' writeDataTable(wb, 3, iris)
#' 
#' saveWorkbook(wb, file = "addFilterExample.xlsx", overwrite = TRUE)
addFilter <- function(wb, sheet, rows, cols){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  sheet <- wb$validateSheet(sheet)
  
  if(length(rows) != 1)
    stop("row must be a numeric of length 1.")
  
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)
  
  wb$worksheets[[sheet]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', paste(getCellRefs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":"))
  
  invisible(wb)
  
}


#' @name removeFilter
#' @title removes worksheet filter from addFilter and writeData
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @export
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#' addWorksheet(wb, "Sheet 3")
#' 
#' writeData(wb, 1, iris)
#' addFilter(wb, 1, row = 1, cols = 1:ncol(iris))
#' 
#' ## Equivalently
#' writeData(wb, 2, x = iris, filter = TRUE)
#' 
#' ## Similarly
#' writeDataTable(wb, 3, iris)
#' 
#' ## remove filters
#' removeFilter(wb, 1:2) ## remove filters
#' removeFilter(wb, 3) ## Does not affect tables!
#' 
#' saveWorkbook(wb, file = "removeFilterExample.xlsx", overwrite = TRUE)
removeFilter <- function(wb, sheet){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  for(s in sheet){
    s <- wb$validateSheet(s)
    wb$worksheets[[s]]$autoFilter <- NULL  
  }
  
  invisible(wb)
  
}


