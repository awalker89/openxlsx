
#' @name createWorkbook
#' @title Create a new Workbook object
#' @description Create a new Workbook object
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
#' saveWorkbook(wb, file = "createWorkbookExample.xlsx", overwrite = TRUE)
createWorkbook <- function(creator = ifelse(.Platform$OS.type == "windows", Sys.getenv("USERNAME"), Sys.getenv("USER"))){
  
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
#' @description save a Workbook object to file
#' @author Alexander Walker
#' @param wb A Workbook object to write to file
#' @param file A character string naming an xlsx file
#' @param overwrite If \code{TRUE}, overwrite any existing file.
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
  exSciPen <- getOption("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  #   if(!grepl("\\.xlsx", file))
  #     file <- paste0(file, ".xlsx")
  
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
#' @description Merge cells within a worksheet
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



#' @name int2col
#' @title Convert integer to Excel column
#' @description Converts an integer to an Excel column label. 
#' @param x A numeric vector
#' @export
#' @examples
#' int2col(1:10)
int2col <- function(x){
  
  if(!is.numeric(x))
    stop("x must be numeric.")
  
  .Call('openxlsx_convert_to_excel_ref', PACKAGE = 'openxlsx', x, LETTERS)
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
#' @description DEPRECATED. Use names().
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
  
  nms <- wb$sheet_names
  nms <- replaceXMLEntities(nms)
  
  return(nms)
}



#' @name addWorksheet
#' @title Add a worksheet to a workbook
#' @description Add a worksheet to a Workbook object
#' @author Alexander Walker
#' @param wb A Workbook object to attach the new worksheet
#' @param sheetName A name for the new worksheet
#' @param gridLines A logical. If \code{FALSE}, the worksheet grid lines will be hidden.
#' @param tabColour Colour of the worksheet tab. A valid colour (belonging to colours()) or a valid hex colour beginning with "#"
#' @param zoom A numeric betwettn 10 and 400. Worksheet zoom level as a percentage.
#' @param header document header. Character vector of length 3 corresponding to positons left, center, right. Use NA to skip a positon.
#' @param footer document footer. Character vector of length 3 corresponding to positons left, center, right. Use NA to skip a positon.
#' @param evenHeader document header for even pages.
#' @param evenFooter document footer for even pages.
#' @param firstHeader document header for first page only.
#' @param firstFooter document footer for first page only.
#' @param visible If FALSE, sheet is hidden else visible.
#' @param paperSize An integer corresponding to a paper size. See ?pageSetup for details.
#' @param orientation One of "portrait" or "landscape"
#' @param hdpi Horizontal DPI. Can be set with options("openxlsx.dpi" = X) or options("openxlsx.hdpi" = X)
#' @param vdpi Vertical DPI. Can be set with options("openxlsx.dpi" = X) or options("openxlsx.vdpi" = X)
#' @details Headers and footers can contain special tags
#' \itemize{
#'   \item{\bold{&[Page]}}{ Page number}
#'   \item{\bold{&[Pages]}}{ Number of pages}
#'   \item{\bold{&[Date]}}{ Current date}
#'   \item{\bold{&[Time]}}{ Current time}
#'   \item{\bold{&[Path]}}{ File path}
#'   \item{\bold{&[File]}}{ File name}
#'   \item{\bold{&[Tab]}}{ Worksheet name}
#' }
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
#' ## Headers and Footers
#' addWorksheet(wb, "Sheet 5",
#' header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#' footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#' evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#' evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#' firstHeader = c("TOP", "OF FIRST", "PAGE"),
#' firstFooter = c("BOTTOM", "OF FIRST", "PAGE"))
#' 
#' addWorksheet(wb, "Sheet 6",
#'              header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'              footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'              firstHeader = c(NA, "Center Header of First Page", NA),
#'              firstFooter = c(NA, "Center Footer of First Page", NA))
#' 
#' addWorksheet(wb, "Sheet 7",
#'              header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'              footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2"))
#' 
#' addWorksheet(wb, "Sheet 8",
#'              firstHeader = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'              firstFooter = c("FIRST ONLY L", NA, "FIRST ONLY R"))
#' 
#' ## Need data on worksheet to see all headers and footers
#' writeData(wb, sheet = 5, 1:400)
#' writeData(wb, sheet = 6, 1:400)
#' writeData(wb, sheet = 7, 1:400)
#' writeData(wb, sheet = 8, 1:400)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "addWorksheetExample.xlsx", overwrite = TRUE)
addWorksheet <- function(wb, sheetName,
                         gridLines = TRUE,
                         tabColour = NULL,
                         zoom = 100,
                         header = NULL,
                         footer = NULL,
                         evenHeader = NULL,
                         evenFooter = NULL,
                         firstHeader = NULL,
                         firstFooter = NULL,
                         visible = TRUE,
                         paperSize = getOption("openxlsx.paperSize", default = 9),
                         orientation = getOption("openxlsx.orientation", default = "portrait"),
                         vdpi = getOption("openxlsx.vdpi", default = getOption("openxlsx.dpi", default = 300)),
                         hdpi = getOption("openxlsx.hdpi", default = getOption("openxlsx.dpi", default = 300))){
  
  
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(tolower(sheetName) %in% tolower(wb$sheet_names))
    stop("A worksheet by that name already exists! Sheet names must be unique case-insensitive.")
  
  if(!is.logical(gridLines) | length(gridLines) > 1)
    stop("gridLines must be a logical of length 1.")
  
  if(nchar(sheetName) > 31)
    stop("sheetName too long! Max length is 31 characters.")
  
  if(!is.null(tabColour))
    tabColour <- validateColour(tabColour, "Invalid tabColour in addWorksheet.")
  
  if(!is.numeric(zoom))
    stop("zoom must be numeric")
  
  if(!is.character(sheetName))
    sheetName <- as.character(sheetName)
  
  if(!is.null(header) & length(header) != 3)
    stop("header must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(footer) & length(footer) != 3)
    stop("footer must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(evenHeader) & length(evenHeader) != 3)
    stop("evenHeader must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(evenFooter) & length(evenFooter) != 3)
    stop("evenFooter must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(firstHeader) & length(firstHeader) != 3)
    stop("firstHeader must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(firstFooter) & length(firstFooter) != 3)
    stop("firstFooter must have length 3 where elements correspond to positions: left, center, right.")
  
  visible <- tolower(visible[1])
  if(!visible %in% c("true",  "false", "hidden", "visible", "veryhidden"))
    stop("visible must be one of: TRUE, FALSE, 'hidden', 'visible', 'veryHidden'")
  
  orientation <- tolower(orientation)
  if(!orientation %in% c("portrait", "landscape"))
    stop("orientation must be 'portrait' or 'landscape'.")
  
  vdpi <- as.integer(vdpi)
  if(is.na(vdpi))
    stop("vdpi must be numeric")
  
  hdpi <- as.integer(hdpi)
  if(is.na(hdpi))
    stop("hdpi must be numeric")
  
  
  
  ## Invalid XML characters
  sheetName <- replaceIllegalCharacters(sheetName)
  
  invisible(wb$addWorksheet(sheetName = sheetName, 
                            showGridLines = gridLines,
                            tabColour = tabColour,
                            zoom = zoom[1],
                            oddHeader = headerFooterSub(header),
                            oddFooter = headerFooterSub(footer),
                            evenHeader = headerFooterSub(evenHeader),
                            evenFooter = headerFooterSub(evenFooter),
                            firstHeader = headerFooterSub(firstHeader),
                            firstFooter = headerFooterSub(firstFooter),
                            visible = visible,
                            paperSize = paperSize,
                            orientation = orientation,
                            vdpi = vdpi,
                            hdpi = hdpi))
} 


#' @name renameWorksheet
#' @title Rename a worksheet
#' @description Rename a worksheet
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
#' @title Convert excel column name to integer index
#' @description Convert excel column name to integer index e.g. "J" to 10
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
  exSciPen <- getOption("scipen")
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
#'   \item{\bold{TEXT}}
#'   \item{\bold{COMMA}{  for comma seperated thousands}}
#'   \item{For date/datetime styling a combination of d, m, y and punctuation marks}
#'   \item{For numeric rouding use "0.00" with the preferred number of decimal places}
#' }
#' 
#' @param border Cell border. A vector of "top", "bottom", "left", "right" or a single string).
#' \itemize{
#'    \item{\bold{"top"}}{ Top border}
#'    \item{\bold{bottom}}{ Bottom border}
#'    \item{\bold{left}}{ Left border}
#'    \item{\bold{right}}{ Right border}
#'    \item{\bold{TopBottom} or \bold{c("top", "bottom")}}{ Top and bottom border}
#'    \item{\bold{LeftRight} or \bold{c("left", "right")}}{ Left and right border}
#'    \item{\bold{TopLeftRight} or \bold{c("top", "left", "right")}}{ Top, Left and right border}
#'    \item{\bold{TopBottomLeftRight} or \bold{c("top", "bottom", "left", "right")}}{ All borders}
#'   }
#'   
#' @param borderColour Colour of cell border vector the same length as the number of sides specified in "border" 
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#"  
#' 
#' @param borderStyle Border line style vector the same length as the number of sides specified in "border"
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
#' A valid colour (belonging to colours()) or a valid hex colour beginning with "#". 
#' --  \bold{Use for conditional formatting styles only.}
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
#' @param wrapText Logical. If \code{TRUE} cell contents will wrap to fit in column.  
#' @param textRotation Rotation of text in degrees. 255 for vertial text.
#' @param indent Horizontal indentation of cell contents.
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
                        textRotation = NULL,
                        indent = NULL){
  
  ### Error checking
  
  ## if num fmt is made up of dd, mm, yy
  numFmt_original <- numFmt[[1]]
  numFmt <- tolower(numFmt_original)
  validNumFmt <- c("general", "number", "currency", "accounting", "date", "longdate", "time", "percentage", "scientific", "text", "3", "4", "comma")
  
  if(numFmt == "date"){
    numFmt <- getOption("openxlsx.dateFormat", getOption("openxlsx.dateformat", "date"))
  }else if(numFmt == "longdate"){
    numFmt <- getOption("openxlsx.datetimeFormat", getOption("openxlsx.datetimeformat", getOption("openxlsx.dateTimeFormat", "longdate")))  
  }else if(!numFmt %in% validNumFmt){
    numFmt <- replaceIllegalCharacters(numFmt_original)
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
                        list("numFmtId" = 49),  # TEXT
                        
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
  
  if(!is.null(indent)){
    if(!is.numeric(indent) & !is.integer(indent))
      stop("indent must be numeric")
  }
  
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
  
  if(!is.null(fontName))
    style$fontName <- list("val" = fontName)
  
  if(!is.null(fontSize))
    style$fontSize <- list("val" = fontSize)
  
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
    border <- paste(border, collapse = "")
    
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
  if(!is.null(halign))
    style$halign <- halign
  
  if(!is.null(valign))
    style$valign <- valign
  
  if(!is.null(indent))
    style$indent <- indent
  
  if(wrapText)
    style$wrapText <- TRUE
  
  if(!is.null(textRotation)){
    if(!is.numeric(textRotation))
      stop("textRotation must be numeric.")
    
    if(textRotation < 0 & textRotation >= -90) {
      textRotation <- (textRotation * -1) + 90
    }
    
    style$textRotation <- round(textRotation[[1]], 0)
  }
  
  if(numFmt != "general"){
    if(numFmt %in% validNumFmt){
      style$numFmt <- numFmtMapping[[numFmt[[1]]]]
    }else{
      style$numFmt <- list("numFmtId" = 9999, formatCode = numFmt)  ## Custom numFmt
    }
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
#' @param gridExpand If \code{TRUE}, style will be applied to all combinations of rows and cols.
#' @param stack If \code{TRUE} the new style is merged with any existing cell styles.  If FALSE, any 
#' existing style is replaced by the new style.
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
addStyle <- function(wb, sheet, style, rows, cols, gridExpand = FALSE, stack = FALSE){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!"Style" %in% class(style))
    stop("style argument must be a Style object.")
  
  if(!is.logical(stack))
    stop("stack parameter must be a logical!")
  
  if(length(cols) == 0 | length(rows) == 0)
    return(invisible(0))
  
  cols <- convertFromExcelRef(cols)
  rows <- as.integer(rows)
  
  ## rows and cols need to be the same length
  if(gridExpand){
    
    n <- length(cols)
    cols <- rep.int(cols, times = length(rows))
    rows <- rep(rows, each = n)
    
  }else if(length(rows) == 1 & length(cols) > 1){
    rows <- rep.int(rows, times = length(cols))
    
  }else if(length(cols) == 1 & length(rows) > 1){
    cols <- rep.int(cols, times = length(rows))
    
  }else if(length(rows) != length(cols)){
    stop("Length of rows and cols must be equal.")
  }
  
  
  wb$addStyle(sheet = sheet, style = style, rows = rows, cols = cols, stack = stack)
  
}


#' @name getCellRefs
#' @title Return excel cell coordinates from (x,y) coordinates
#' @description Return excel cell coordinates from (x,y) coordinates
#' @author Alexander Walker
#' @param cellCoords A data.frame with two columns coordinate pairs. 
#' @return Excel alphanumeric cell reference
getCellRefs <- function(cellCoords){
  l <- convert_to_excel_ref(cols = unlist(cellCoords[,2]), LETTERS = LETTERS)
  paste0(l, cellCoords[,1])
}











#' @name freezePane
#' @title Freeze a worksheet pane
#' @description Freeze a worksheet pane 
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param firstActiveRow Top row of active region
#' @param firstActiveCol Furthest left column of active region
#' @param firstRow If \code{TRUE}, freezes the first row (equivalent to firstActiveRow = 2)
#' @param firstCol If \code{TRUE}, freezes the first column (equivalent to firstActiveCol = 2)
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
  
  
  if(is.null(firstActiveRow) & is.null(firstActiveCol) & !firstRow & !firstCol)
    return(invisible(0))
  
  if(!is.logical(firstRow)) 
    stop("firstRow must be TRUE/FALSE")
  
  if(!is.logical(firstCol)) 
    stop("firstCol must be TRUE/FALSE")
  
  
  
  
  if(firstRow & !firstCol){
    invisible(wb$freezePanes(sheet, firstRow = firstRow))
  }else if(firstCol & !firstRow){
    invisible(wb$freezePanes(sheet, firstCol = firstCol))
  }else if(firstRow & firstCol){
    invisible(wb$freezePanes(sheet, firstActiveRow = 2L, firstActiveCol = 2L))
  }else{ ## else both firstRow and firstCol are FALSE
    
    ## Convert to numeric if column letter given
    if(!is.null(firstActiveRow)){
      firstActiveRow <- convertFromExcelRef(firstActiveRow)
    }else{
      firstActiveRow <- 1L
    }
    
    if(!is.null(firstActiveCol)){
      firstActiveCol <- convertFromExcelRef(firstActiveCol)
    }else{
      firstActiveCol <- 1L
    }
    
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
#' @description Insert an image into a worksheet
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param file An image file. Valid file types are: jpeg, png, bmp
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
#' @description Set worksheet row heights
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
  
  
  heights <- as.character(as.numeric(heights))
  names(heights) <- rows
  
  wb$setRowHeights(sheet, rows, heights)
  
}

#' @name setColWidths
#' @title Set worksheet column widths
#' @description Set worksheet column widths to specific width or "auto".
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of cols to set width
#' @param widths widths to set cols to specified in Excel column width units or "auto" for automatic sizing. The widths argument is
#' recycled to the length of cols.
#' @param hidden Logical vector. If TRUE the column is hidden.
#' @param ignoreMergedCells Ignore any cells that have been merged with other cells in the calculation of "auto" column widths.
#' @details The global min and max column width for "auto" columns is set by (default values show):
#' \itemize{
#'   \item{options("openxlsx.minWidth" = 3)}
#'   \item{options("openxlsx.maxWidth" = 250)} ## This is the maximum width allowed in Excel
#' }
#' 
#' NOTE: The calculation of column widths can be slow for large worksheets.
#' 
#' @seealso \code{\link{removeColWidths}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- createWorkbook()
#' 
#' ## Add a worksheet
#' addWorksheet(wb, "Sheet 1") 
#'
#'
#' ## set col widths
#' setColWidths(wb, 1, cols = c(1,4,6,7,9), widths = c(16,15,12,18,33))
#'
#' ## auto columns
#' addWorksheet(wb, "Sheet 2")
#' writeData(wb, sheet = 2, x = iris)
#' setColWidths(wb, sheet = 2, cols = 1:5, widths = "auto")
#'   
#' ## Save workbook
#' saveWorkbook(wb, "setColWidthsExample.xlsx", overwrite = TRUE)
setColWidths <- function(wb, sheet, cols, widths = 8.43, hidden = rep(FALSE, length(cols)), ignoreMergedCells = FALSE){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  widths <- tolower(widths)  ## possibly "auto"
  if(ignoreMergedCells)
    widths[widths == "auto"] <- "auto2"
  
  if(length(widths) > length(cols))
    stop("More widths than columns supplied.")
  
  if(length(hidden) > length(cols))
    stop("hidden argument is longer than cols.")
  
  if(length(widths) < length(cols))
    widths <- rep(widths, length.out = length(cols))
  
  if(length(hidden) < length(cols))
    hidden <- rep(hidden, length.out = length(cols))
  
  ## Remove duplicates
  widths <- widths[!duplicated(cols)]
  hidden <- hidden[!duplicated(cols)]
  cols <- cols[!duplicated(cols)]
  cols <- convertFromExcelRef(cols)
  
  if(length(wb$colWidths[[sheet]]) > 0){
    
    existing_cols <- names(wb$colWidths[[sheet]])
    existing_widths <- unname(wb$colWidths[[sheet]])
    existing_hidden <- attr(wb$colWidths[[sheet]], "hidden")
    
    ## check for existing custom widths
    flag <- existing_cols %in% cols
    if(any(flag)){
      existing_cols <- existing_cols[!flag]
      existing_widths <- existing_widths[!flag]
      existing_hidden <- existing_hidden[!flag]
    }

    all_names <- c(existing_cols, cols)
    all_widths <- c(existing_widths, widths)
    all_hidden <- c(existing_hidden, as.character(as.integer(hidden)))
    
    ord <- order(as.integer(all_names))
    all_names <- all_names[ord]
    all_widths <- all_widths[ord]
    all_hidden <- all_hidden[ord]
    
    
    names(all_widths) <- all_names
    wb$colWidths[[sheet]] <- all_widths
    attr(wb$colWidths[[sheet]], "hidden") <- all_hidden
    
    
    
  }else{
    
    names(widths) <- cols
    wb$colWidths[[sheet]] <- widths
    attr(wb$colWidths[[sheet]], "hidden") <- as.character(as.integer(hidden))
    
  }
  
  
  invisible(0)
}


#' @name removeColWidths
#' @title Remove column widths from a worksheet
#' @description Remove column widths from a worksheet
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Indices of colunss to remove custom width (if any) from.
#' @seealso \code{\link{setColWidths}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
#' 
#' ## remove column widths in columns 1 to 20
#' removeColWidths(wb, 1, cols = 1:20)
#' saveWorkbook(wb, "removeColWidthsExample.xlsx", overwrite = TRUE)
removeColWidths <- function(wb, sheet, cols){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)
  
  customCols <- as.integer(names(wb$colWidths[[sheet]]))
  removeInds <- which(customCols %in% cols)
  if(length(removeInds) > 0){
    
    remainingCols <- customCols[-removeInds]
    if(length(remainingCols) == 0){
      wb$colWidths[[sheet]] <- list()
    }else{
      
      rem_widths <- wb$colWidths[[sheet]][-removeInds]
      names(rem_widths) <- as.character(remainingCols)
      wb$colWidths[[sheet]] <- rem_widths
      
    }
    
    
    
  }
  
  
}



#' @name removeRowHeights
#' @title Remove custom row heights from a worksheet
#' @description Remove row heights from a worksheet
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Indices of rows to remove custom height (if any) from.
#' @seealso \code{\link{setRowHeights}}
#' @export
#' @examples
#' ## Create a new workbook
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
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
    dev.copy(tiff, filename = fileName, width = width, height = height, units = units, compression = "none", res = dpi)
  }
  
  ## write image
  invisible(dev.off())
  
  insertImage(wb = wb, sheet = sheet, file = fileName, width = width, height = height, startRow = startRow, startCol = startCol, units = units, dpi = dpi)
  
}



#' @name replaceStyle
#' @title Replace an existing cell style
#' @description Replace an existing cell style
#' @author Alexander Walker
#' @param wb A workbook object
#' @param index Index of style object to replace
#' @param newStyle A style to replace the exising style as position index
#' @description Replace a style object
#' @export
#' @seealso \code{\link{getStyles}}
#' @examples
#' 
#' ## load a workbook 
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
#' 
#' ## create a new style and replace style 2
#' 
#' newStyle <- createStyle(fgFill = "#00FF00")
#'  
#' ## replace style 2
#' getStyles(wb)[1:3] ## prints styles
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
#' @description Returns list of style objects in the workbook
#' @param wb A workbook object
#' @export
#' @seealso \code{\link{replaceStyle}}
#' @examples
#' ## load a workbook 
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
#' getStyles(wb)[1:3]
getStyles <- function(wb){
  
  nStyles <- length(wb$styleObjects)
  
  if(nStyles == 0)
    stop("Workbook has no existing styles.")
  
  styles <- lapply(wb$styleObjects, "[[", "style")
  
  return(styles)
}



#' @name removeWorksheet
#' @title Remove a worksheet from a workbook
#' @description Remove a worksheet from a Workbook object
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @description Remove a worksheet from a workbook
#' @export
#' @examples
#' ## load a workbook 
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
#' 
#' ## Remove sheet 2
#' removeWorksheet(wb, 2)
#' 
#' ## save the modified workbook
#' saveWorkbook(wb, "removeWorksheetExample.xlsx", overwrite = TRUE)
removeWorksheet <- function(wb, sheet){
  
  if(class(wb) != "Workbook")
    stop("wb must be a Workbook object!")
  
  if(length(sheet) != 1)
    stop("sheet must have length 1.")
  
  wb$deleteWorksheet(sheet)
  
  invisible(0)
}


#' @name deleteData
#' @title Delete cell data
#' @description Delete contents and styling from a cell.
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Rows to delete data from.
#' @param cols columns to delete data from.
#' @param gridExpand If \code{TRUE}, all data in rectangle min(rows):max(rows) X min(cols):max(cols)
#' will be removed.
#' @export
#' @examples
#' ## write some data
#' wb <- createWorkbook()
#' addWorksheet(wb, "Worksheet 1")
#' x <- data.frame(matrix(runif(200), ncol = 10)) 
#' writeData(wb, sheet = 1, x = x, startCol = 2, startRow = 3, colNames = FALSE)
#' 
#' ## delete some data
#' deleteData(wb, sheet = 1, cols = 3:5, rows = 5:7, gridExpand = TRUE)
#' deleteData(wb, sheet = 1, cols = 7:9, rows = 5:7, gridExpand = TRUE)
#' deleteData(wb, sheet = 1, cols = LETTERS, rows = 18, gridExpand = TRUE)
#' 
#' saveWorkbook(wb, "deleteDataExample.xlsx", overwrite = TRUE)
deleteData <- function(wb, sheet, cols, rows, gridExpand = FALSE){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  
  wb$worksheets[[sheet]]$sheet_data$delete(rows_in = rows, cols_in = cols, grid_expand = gridExpand)
  
  
  invisible(0)
}


#' @name modifyBaseFont
#' @title Modify the default font
#' @description Modify the default font for this workbook
#' @author Alexander Walker
#' @param wb A workbook object
#' @param fontSize font size
#' @param fontColour font colour
#' @param fontName Name of a font
#' @details The font name is not validated in anyway.  Excel replaces unknown font names
#' with Arial. Base font is black, size 11, Calibri.
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
modifyBaseFont <- function(wb, fontSize = 11, fontColour = "black", fontName = "Calibri"){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  
  if(fontSize < 0) stop("Invalid fontSize")
  fontColour <- validateColour(fontColour)
  
  wb$styles$fonts[[1]] <- sprintf('<font><sz val="%s"/><color rgb="%s"/><name val="%s"/></font>', fontSize, fontColour, fontName)
  
}


#' @name getBaseFont
#' @title Return the workbook default font
#' @description Return the workbook default font
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


#' @name setHeaderFooter
#' @title Set document headers and footers
#' @description Set document headers and footers
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param header document header. Character vector of length 3 corresponding to positons left, center, right. Use NA to skip a positon.
#' @param footer document footer. Character vector of length 3 corresponding to positons left, center, right. Use NA to skip a positon.
#' @param evenHeader document header for even pages.
#' @param evenFooter document footer for even pages.
#' @param firstHeader document header for first page only.
#' @param firstFooter document footer for first page only.
#' @details Headers and footers can contain special tags
#' \itemize{
#'   \item{\bold{&[Page]}}{ Page number}
#'   \item{\bold{&[Pages]}}{ Number of pages}
#'   \item{\bold{&[Date]}}{ Current date}
#'   \item{\bold{&[Time]}}{ Current time}
#'   \item{\bold{&[Path]}}{ File path}
#'   \item{\bold{&[File]}}{ File name}
#'   \item{\bold{&[Tab]}}{ Worksheet name}
#' }
#' @export
#' @seealso \code{\link{addWorksheet}} to set headers and footers when adding a worksheet
#' @examples
#' wb <- createWorkbook()
#' 
#' addWorksheet(wb, "S1")
#' addWorksheet(wb, "S2")
#' addWorksheet(wb, "S3")
#' addWorksheet(wb, "S4")
#' 
#' writeData(wb, 1, 1:400)
#' writeData(wb, 2, 1:400)
#' writeData(wb, 3, 3:400)
#' writeData(wb, 4, 3:400)
#' 
#' setHeaderFooter(wb, sheet = "S1",  
#'                 header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
#'                 footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
#'                 evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
#'                 evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
#'                 firstHeader = c("TOP", "OF FIRST", "PAGE"),
#'                 firstFooter = c("BOTTOM", "OF FIRST", "PAGE"))
#' 
#' setHeaderFooter(wb, sheet = 2,  
#'                 header = c("&[Date]", "ALL HEAD CENTER 2", "&[Page] / &[Pages]"),
#'                 footer = c("&[Path]&[File]", NA, "&[Tab]"),
#'                 firstHeader = c(NA, "Center Header of First Page", NA),
#'                 firstFooter = c(NA, "Center Footer of First Page", NA))
#' 
#' setHeaderFooter(wb, sheet = 3,  
#'                 header = c("ALL HEAD LEFT 2", "ALL HEAD CENTER 2", "ALL HEAD RIGHT 2"),
#'                 footer = c("ALL FOOT RIGHT 2", "ALL FOOT CENTER 2", "ALL FOOT RIGHT 2"))
#' 
#' setHeaderFooter(wb, sheet = 4,  
#'                 firstHeader = c("FIRST ONLY L", NA, "FIRST ONLY R"),
#'                 firstFooter = c("FIRST ONLY L", NA, "FIRST ONLY R"))
#' 
#' 
#' saveWorkbook(wb, "setHeaderFooterExample.xlsx", overwrite = TRUE)
setHeaderFooter <- function(wb, sheet,
                            header = NULL,
                            footer = NULL,
                            evenHeader = NULL,
                            evenFooter = NULL,
                            firstHeader = NULL,
                            firstFooter = NULL){
  
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  sheet <- wb$validateSheet(sheet)
  
  if(!is.null(header) & length(header) != 3)
    stop("header must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(footer) & length(footer) != 3)
    stop("footer must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(evenHeader) & length(evenHeader) != 3)
    stop("evenHeader must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(evenFooter) & length(evenFooter) != 3)
    stop("evenFooter must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(firstHeader) & length(firstHeader) != 3)
    stop("firstHeader must have length 3 where elements correspond to positions: left, center, right.")
  
  if(!is.null(firstFooter) & length(firstFooter) != 3)
    stop("firstFooter must have length 3 where elements correspond to positions: left, center, right.")
  
  oddHeader = headerFooterSub(header)
  oddFooter = headerFooterSub(footer)
  evenHeader = headerFooterSub(evenHeader)
  evenFooter = headerFooterSub(evenFooter)
  firstHeader = headerFooterSub(firstHeader)
  firstFooter = headerFooterSub(firstFooter)
  
  naToNULLList <- function(x){
    lapply(x, function(x) {
      if(is.na(x))
        return(NULL)
      x})
  }
  
  hf <- list(oddHeader = naToNULLList(oddHeader),
             oddFooter = naToNULLList(oddFooter),
             evenHeader = naToNULLList(evenHeader),
             evenFooter = naToNULLList(evenFooter), 
             firstHeader = naToNULLList(firstHeader),
             firstFooter = naToNULLList(firstFooter))
  
  if(all(sapply(hf, length) == 0))
    hf <- NULL
  
  
  wb$worksheets[[sheet]]$headerFooter <- hf
  
  
}




#' @name pageSetup
#' @title Set page margins, orientation and print scaling
#' @description Set page margins, orientation and print scaling
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
#' @param fitToWidth If \code{TRUE}, worksheet is scaled to fit to page width on printing.
#' @param fitToHeight If \code{TRUE}, worksheet is scaled to fit to page height on printing.
#' @param paperSize See details. Default value is 9 (A4 paper).
#' @param printTitleRows Rows to repeat at top of page when printing. Integer vector.
#' @param printTitleCols Columns to repeat at left when printing. Integer vector.
#' @export
#' @details
#' paperSize is an integer corresponding to: 
#' \itemize{
#' \item{\bold{1}}{ Letter paper (8.5 in. by 11 in.)} 
#' \item{\bold{2}}{ Letter small paper (8.5 in. by 11 in.)} 
#' \item{\bold{3}}{ Tabloid paper (11 in. by 17 in.)} 
#' \item{\bold{4}}{ Ledger paper (17 in. by 11 in.)} 
#' \item{\bold{5}}{ Legal paper (8.5 in. by 14 in.)} 
#' \item{\bold{6}}{ Statement paper (5.5 in. by 8.5 in.)} 
#' \item{\bold{7}}{ Executive paper (7.25 in. by 10.5 in.)} 
#' \item{\bold{8}}{ A3 paper (297 mm by 420 mm)} 
#' \item{\bold{9}}{ A4 paper (210 mm by 297 mm)} 
#' \item{\bold{10}}{ A4 small paper (210 mm by 297 mm)} 
#' \item{\bold{11}}{ A5 paper (148 mm by 210 mm)} 
#' \item{\bold{12}}{ B4 paper (250 mm by 353 mm)} 
#' \item{\bold{13}}{ B5 paper (176 mm by 250 mm)} 
#' \item{\bold{14}}{ Folio paper (8.5 in. by 13 in.)} 
#' \item{\bold{15}}{ Quarto paper (215 mm by 275 mm)} 
#' \item{\bold{16}}{ Standard paper (10 in. by 14 in.)} 
#' \item{\bold{17}}{ Standard paper (11 in. by 17 in.)} 
#' \item{\bold{18}}{ Note paper (8.5 in. by 11 in.)} 
#' \item{\bold{19}}{ #9 envelope (3.875 in. by 8.875 in.)} 
#' \item{\bold{20}}{ #10 envelope (4.125 in. by 9.5 in.)} 
#' \item{\bold{21}}{ #11 envelope (4.5 in. by 10.375 in.)} 
#' \item{\bold{22}}{ #12 envelope (4.75 in. by 11 in.)} 
#' \item{\bold{23}}{ #14 envelope (5 in. by 11.5 in.)} 
#' \item{\bold{24}}{ C paper (17 in. by 22 in.)} 
#' \item{\bold{25}}{ D paper (22 in. by 34 in.)} 
#' \item{\bold{26}}{ E paper (34 in. by 44 in.)} 
#' \item{\bold{27}}{ DL envelope (110 mm by 220 mm)} 
#' \item{\bold{28}}{ C5 envelope (162 mm by 229 mm)} 
#' \item{\bold{29}}{ C3 envelope (324 mm by 458 mm)} 
#' \item{\bold{30}}{ C4 envelope (229 mm by 324 mm)} 
#' \item{\bold{31}}{ C6 envelope (114 mm by 162 mm)} 
#' \item{\bold{32}}{ C65 envelope (114 mm by 229 mm)} 
#' \item{\bold{33}}{ B4 envelope (250 mm by 353 mm)} 
#' \item{\bold{34}}{ B5 envelope (176 mm by 250 mm)} 
#' \item{\bold{35}}{ B6 envelope (176 mm by 125 mm)} 
#' \item{\bold{36}}{ Italy envelope (110 mm by 230 mm)} 
#' \item{\bold{37}}{ Monarch envelope (3.875 in. by 7.5 in.).} 
#' \item{\bold{38}}{ 6 3/4 envelope (3.625 in. by 6.5 in.)} 
#' \item{\bold{39}}{ US standard fanfold (14.875 in. by 11 in.)} 
#' \item{\bold{40}}{ German standard fanfold (8.5 in. by 12 in.)} 
#' \item{\bold{41}}{ German legal fanfold (8.5 in. by 13 in.)} 
#' \item{\bold{42}}{ ISO B4 (250 mm by 353 mm)} 
#' \item{\bold{43}}{ Japanese double postcard (200 mm by 148 mm)} 
#' \item{\bold{44}}{ Standard paper (9 in. by 11 in.)} 
#' \item{\bold{45}}{ Standard paper (10 in. by 11 in.)} 
#' \item{\bold{46}}{ Standard paper (15 in. by 11 in.)} 
#' \item{\bold{47}}{ Invite envelope (220 mm by 220 mm)} 
#' \item{\bold{50}}{ Letter extra paper (9.275 in. by 12 in.)} 
#' \item{\bold{51}}{ Legal extra paper (9.275 in. by 15 in.)} 
#' \item{\bold{52}}{ Tabloid extra paper (11.69 in. by 18 in.)} 
#' \item{\bold{53}}{ A4 extra paper (236 mm by 322 mm)} 
#' \item{\bold{54}}{ Letter transverse paper (8.275 in. by 11 in.)} 
#' \item{\bold{55}}{ A4 transverse paper (210 mm by 297 mm)} 
#' \item{\bold{56}}{ Letter extra transverse paper (9.275 in. by 12 in.)} 
#' \item{\bold{57}}{ SuperA/SuperA/A4 paper (227 mm by 356 mm)} 
#' \item{\bold{58}}{ SuperB/SuperB/A3 paper (305 mm by 487 mm)} 
#' \item{\bold{59}}{ Letter plus paper (8.5 in. by 12.69 in.)} 
#' \item{\bold{60}}{ A4 plus paper (210 mm by 330 mm)} 
#' \item{\bold{61}}{ A5 transverse paper (148 mm by 210 mm)} 
#' \item{\bold{62}}{ JIS B5 transverse paper (182 mm by 257 mm)} 
#' \item{\bold{63}}{ A3 extra paper (322 mm by 445 mm)} 
#' \item{\bold{64}}{ A5 extra paper (174 mm by 235 mm)} 
#' \item{\bold{65}}{ ISO B5 extra paper (201 mm by 276 mm)} 
#' \item{\bold{66}}{ A2 paper (420 mm by 594 mm)} 
#' \item{\bold{67}}{ A3 transverse paper (297 mm by 420 mm)} 
#' \item{\bold{68}}{ A3 extra transverse paper (322 mm by 445 mm)}
#' }
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
#' 
#' ## print titles
#' addWorksheet(wb, "print_title_rows")
#' addWorksheet(wb, "print_title_cols")
#'
#' writeData(wb, "print_title_rows", rbind(iris, iris, iris, iris))
#' writeData(wb, "print_title_cols", x = rbind(mtcars, mtcars, mtcars), rowNames = TRUE)
#' 
#' pageSetup(wb, sheet = "print_title_rows", printTitleRows = 1) ## first row
#' pageSetup(wb, sheet = "print_title_cols", printTitleCols = 1, printTitleRows = 1)
#' 
#' 
#' saveWorkbook(wb, "pageSetupExample.xlsx", overwrite = TRUE)
pageSetup <- function(wb, sheet, orientation = NULL, scale = 100,
                      left = 0.7, right = 0.7, top = 0.75, bottom = 0.75,
                      header = 0.3, footer = 0.3,
                      fitToWidth = FALSE, fitToHeight = FALSE, paperSize = NULL,
                      printTitleRows = NULL, printTitleCols = NULL){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  sheet <- wb$validateSheet(sheet)
  xml <- wb$worksheets[[sheet]]$pageSetup
  
  if(!is.null(orientation)){
    orientation <- tolower(orientation)
    if(!orientation %in% c("portrait", "landscape")) stop("Invalid page orientation.")
  }else{
    orientation <- ifelse(grepl("landscape", xml), "landscape", "portrait") ## get existing
  }
  
  if(scale < 10 | scale > 400)
    stop("Scale must be between 10 and 400.")
  
  if(!is.null(paperSize)){
    paperSizes <- 1:68
    paperSizes <- paperSizes[!paperSizes %in% 48:49]
    if(!paperSize %in% paperSizes)
      stop("paperSize must be an integer in range [1, 68]. See ?pageSetup details.")
    paperSize <- as.integer(paperSize)
  }else{
    paperSize <- regmatches(xml, regexpr('(?<=paperSize=")[0-9]+', xml, perl = TRUE)) ## get existing
  }

  
  ##############################
  ## Keep defaults on orientation, hdpi, vdpi, paperSize
  hdpi <- regmatches(xml, regexpr('(?<=horizontalDpi=")[0-9]+', xml, perl = TRUE))
  vdpi <- regmatches(xml, regexpr('(?<=verticalDpi=")[0-9]+', xml, perl = TRUE))
  
  
  ##############################
  ## Update
  wb$worksheets[[sheet]]$pageSetup <- sprintf('<pageSetup paperSize="%s" orientation="%s" scale = "%s" fitToWidth="%s" fitToHeight="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>', 
                                              paperSize, orientation, scale, as.integer(fitToWidth), as.integer(fitToHeight), hdpi, vdpi)
  
  if(fitToHeight | fitToWidth)
    wb$worksheets[[sheet]]$sheetPr <- unique(c(wb$worksheets[[sheet]]$sheetPr, '<pageSetUpPr fitToPage="1"/>'))
  
  wb$worksheets[[sheet]]$pageMargins <- 
    sprintf('<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>', left, right, top, bottom, header, footer)
  
  ## print Titles
  if(!is.null(printTitleRows) & is.null(printTitleCols)){
    
    if(!is.numeric(printTitleRows))
      stop("printTitleRows must be numeric.")
    
    wb$createNamedRegion(ref1 = paste0("$", min(printTitleRows)),
                         ref2 = paste0("$", max(printTitleRows)),
                         name = "_xlnm.Print_Titles",
                         sheet = names(wb)[[sheet]],
                         localSheetId = sheet - 1L)
    
    
  }else if(!is.null(printTitleCols) & is.null(printTitleRows)){
    
    if(!is.numeric(printTitleCols))
      stop("printTitleCols must be numeric.")
    
    cols <- .Call('openxlsx_convert_to_excel_ref', range(printTitleCols), LETTERS, PACKAGE="openxlsx")
    wb$createNamedRegion(ref1 = paste0("$", cols[1]),
                         ref2 = paste0("$", cols[2]),
                         name = "_xlnm.Print_Titles",
                         sheet = names(wb)[[sheet]],
                         localSheetId = sheet - 1L)
    
    
  }else if(!is.null(printTitleCols) & !is.null(printTitleRows)){
    
    if(!is.numeric(printTitleRows))
      stop("printTitleRows must be numeric.")
    
    if(!is.numeric(printTitleCols))
      stop("printTitleCols must be numeric.")
    
    
    cols <- .Call("openxlsx_convert_to_excel_ref", range(printTitleCols), LETTERS)
    rows <- range(printTitleRows)
    
    cols <- paste(paste0("$", cols[1]), paste0("$", cols[2]), sep = ":")
    rows <- paste(paste0("$", rows[1]), paste0("$", rows[2]), sep = ":")
    localSheetId <- sheet - 1L
    sheet <- names(wb)[[sheet]]
    
    wb$workbook$definedNames <- c(wb$workbook$definedNames, 
                                  sprintf('<definedName name="_xlnm.Print_Titles" localSheetId="%s">\'%s\'!%s,\'%s\'!%s</definedName>', localSheetId, sheet, cols, sheet, rows)
    )
    
  }
  
  
}




#' @name showGridLines
#' @title Set worksheet gridlines to show or hide.
#' @description Set worksheet gridlines to show or hide.
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param showGridLines A logical. If \code{TRUE}, grid lines are hidden.
#' @export
#' @examples
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package = "openxlsx"))
#' names(wb) ## list worksheets in workbook
#' showGridLines(wb, 1, showGridLines = FALSE)
#' showGridLines(wb, "testing", showGridLines = FALSE)
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
#' @description Get/set order of worksheets in a Workbook object
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
  
  wb$sheetOrder
  
}

#' @rdname worksheetOrder
#' @param wb A workbook object
#' @param value Vector specifying order to write worksheets to file
#' @export
`worksheetOrder<-` <- function(wb, value) {
  
  if(!"Workbook" %in% class(wb))
    stop("Argument must be a Workbook.")
  
  if(any(value != as.integer(value)))
    stop("values must be integers")
  
  value <- as.integer(value)
  
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
#' @description Convert from excel date number to R Date type
#' @param x A vector of integers
#' @param origin date. Default value is for Windows Excel 2010
#' @param ... additional parameters passed to as.Date()
#' @details Excel stores dates as number of days from some origin day
#' @seealso \code{\link{writeData}}
#' @export
#' @examples
#' ##2014 April 21st to 25th
#' convertToDate(c(41750, 41751, 41752, 41753, 41754, NA) )
#' convertToDate(c(41750.2, 41751.99, NA, 41753 ))
convertToDate <- function(x, origin = "1900-01-01", ...){
  
  x <- as.numeric(x)
  notNa <- !is.na(x)
  if(origin == "1900-01-01")
    x[notNa] <- x[notNa] - 2
  
  return(as.Date(x, origin = origin, ...))
  
  
}


#' @name convertToDateTime
#' @title Convert from excel time number to R POSIXct type.
#' @description Convert from excel time number to R POSIXct type.
#' @param x A numeric vector
#' @param origin date. Default value is for Windows Excel 2010
#' @param ... Additional parameters passed to as.POSIXct
#' @details Excel stores dates as number of days from some origin date
#' @export
#' @examples
#' ## 2014-07-01, 2014-06-30, 2014-06-29
#' x <- c(41821.8127314815, 41820.8127314815, NA, 41819, NaN) 
#' convertToDateTime(x)
#' convertToDateTime(x, tx = "Australia/Perth")
convertToDateTime <- function(x, origin = "1900-01-01", ...){
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- options("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  x <- as.numeric(x)
  rem <- x %% 1
  date <- convertToDate(x, origin)
  fraction <- 24*rem
  hrs <- floor(fraction)
  minFrac <- (fraction-hrs)*60
  mins <- floor(minFrac)
  secs <- (minFrac - mins)*60
  y <- paste(hrs, mins, secs, sep = ":")
  y <- format(strptime(y, "%H:%M:%S"), "%H:%M:%S") 
  
  notNA <- !is.na(x)
  dateTime = rep(NA, length(x))
  dateTime[notNA] <- as.POSIXct(paste(date[notNA], y[notNA]), ...)
  dateTime = .POSIXct(dateTime)
  
  return(dateTime)
}



#' @name names
#' @title get or set worksheet names
#' @description get or set worksheet names
#' @aliases names.Workbook
#' @export
#' @method names Workbook
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
  nms <- x$sheet_names
  nms <- replaceXMLEntities(nms)
}

#' @rdname names
#' @param value a character vector the same length as wb
#' @export
`names<-.Workbook` <- function(x, value) {
  
  if(any(duplicated(tolower(value))))
    stop("Worksheet names must be unique.")
  
  existing_sheets <- x$sheet_names
  inds <- which(value != existing_sheets)
  
  if(length(inds) == 0)
    return(invisible(x))
  
  if(length(value) != length(x$worksheets))
    stop(sprintf("names vector must have length equal to number of worksheets in Workbook [%s]", length(existing_sheets)))
  
  if(any(nchar(value) > 31)){
    warning("Worksheet names must less than 32 characters. Truncating names...")
    value[nchar(value) > 31] <- sapply(value[nchar(value) > 31], substr, start = 1, stop = 31)
  }
  
  for(i in inds)
    invisible(x$setSheetName(i, value[[i]]))
  
  invisible(x)
  
}



#' @name createNamedRegion
#' @title Create a named region.
#' @description Create a named region
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param rows Numeric vector specifying rows to include in region
#' @param cols Numeric vector specifying columns to include in region
#' @param name Name for region. A character vector of length 1. Note region names musts be case-insensitive unique.
#' @details Region is given by: min(cols):max(cols) X min(rows):max(rows)
#' @export
#' @seealso \code{\link{getNamedRegions}}
#' @examples
#' ## create named regions
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' 
#' ## specify region
#' writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' createNamedRegion(wb = wb,
#'                   sheet = 1,
#'                   name = "iris",
#'                   rows = 1:(nrow(iris)+1),
#'                   cols = 1:ncol(iris))
#' 
#' 
#' ## using writeData 'name' argument
#' writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#' 
#' out_file <- tempfile(fileext = ".xlsx")
#' saveWorkbook(wb, out_file, overwrite = TRUE)
#' 
#' ## see named regions
#' getNamedRegions(wb) ## From Workbook object
#' getNamedRegions(out_file) ## From xlsx file
#' 
#' ## read named regions
#' df <- read.xlsx(wb, namedRegion = "iris")
#' head(df)
#' 
#' df <- read.xlsx(out_file, namedRegion = "iris2")
#' head(df)
createNamedRegion <- function(wb, sheet, cols, rows, name){
  
  sheet <- wb$validateSheet(sheet)
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!is.numeric(rows))
    stop("rows argument must be a numeric/integer vector")
  
  if(!is.numeric(cols))
    stop("cols argument must be a numeric/integer vector")
  
  ## check name doesn't already exist
  ## named region
  
  ex_names <- regmatches(wb$workbook$definedNames, regexpr('(?<=name=")[^"]+', wb$workbook$definedNames, perl = TRUE))
  ex_names <- tolower(replaceXMLEntities(ex_names))
  
  if(tolower(name) %in% ex_names){
    stop(sprintf("Named region with name '%s' already exists!", name))
  }else if(grepl("[^A-Z0-9_\\.]", name[1], ignore.case = TRUE)){
    stop("Invalid characters in name")
  }else if(grepl('^[A-Z]{1,3}[0-9]+$', name)){
    stop("name cannot look like a cell reference.")
  }
  
  
  cols <- round(cols)
  rows <- round(rows)
  
  startCol <- min(cols)
  endCol <- max(cols)
  
  startRow <- min(rows)
  endRow <- max(rows)
  
  ref1 <- paste0("$", .Call("openxlsx_convert_to_excel_ref", startCol, LETTERS, PACKAGE = "openxlsx"), "$", startRow)
  ref2 <- paste0("$", .Call("openxlsx_convert_to_excel_ref", endCol, LETTERS, PACKAGE = "openxlsx"), "$", endRow)
  
  invisible(
    wb$createNamedRegion(ref1 = ref1, ref2 = ref2, name = name, sheet = wb$sheet_names[sheet])
  )
  
}



#' @name getNamedRegions
#' @title Get named regions
#' @description Return a vector of named regions in a xlsx file or
#' Workbook object
#' @param x An xlsx file or Workbook object
#' @export
#' @seealso \code{\link{createNamedRegion}}
#' @examples
#' ## create named regions
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' 
#' ## specify region
#' writeData(wb, sheet = 1, x = iris, startCol = 1, startRow = 1)
#' createNamedRegion(wb = wb,
#'                   sheet = 1,
#'                   name = "iris",
#'                   rows = 1:(nrow(iris)+1),
#'                   cols = 1:ncol(iris))
#' 
#' 
#' ## using writeData 'name' argument to create a named region
#' writeData(wb, sheet = 1, x = iris, name = "iris2", startCol = 10)
#' 
#' out_file <- tempfile(fileext = ".xlsx")
#' saveWorkbook(wb, out_file, overwrite = TRUE)
#' 
#' ## see named regions
#' getNamedRegions(wb) ## From Workbook object
#' getNamedRegions(out_file) ## From xlsx file
#' 
#' ## read named regions
#' df <- read.xlsx(wb, namedRegion = "iris")
#' head(df)
#' 
#' df <- read.xlsx(out_file, namedRegion = "iris2")
#' head(df)
getNamedRegions <- function(x){
  
  UseMethod("getNamedRegions", x) 
  
}

#' @export
getNamedRegions.default <- function(x){
  
  if(!file.exists(x))
    stop(sprintf("File '%s' does not exist.", x))
  
  xmlDir <- file.path(tempdir(), "named_regions_tmp")
  xmlFiles <- unzip(x, exdir = xmlDir)
  
  workbook <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  workbook <- unlist(readLines(workbook, warn = FALSE, encoding = "UTF-8"))
  
  dn <- .Call("openxlsx_getChildlessNode", removeHeadTag(workbook), "<definedName ", PACKAGE = "openxlsx")
  if(length(dn) == 0)
    return(NULL)
  
  dn_names <- get_named_regions_from_string(dn = dn)
  
  unlink(xmlDir, recursive = TRUE, force = TRUE)
  
  return(dn_names)
}


#' @export
getNamedRegions.Workbook <- function(x){
  
  dn <- x$workbook$definedNames
  if(length(dn) == 0)
    return(NULL)
  
  dn_names <- get_named_regions_from_string(dn = dn)
  
  return(dn_names)
  
}






#' @name addFilter
#' @title Add column filters
#' @description Add excel column filters to a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols columns to add filter to. 
#' @param rows A row number.
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
#' writeData(wb, 2, x = iris, withFilter = TRUE)
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
#' @title Remove a worksheet filter
#' @description Removes filters from addFilter() and writeData()
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
#' writeData(wb, 2, x = iris, withFilter = TRUE)
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
    wb$worksheets[[s]]$autoFilter <- character(0)  
  }
  
  invisible(wb)
  
}











#' @name setHeader
#' @title Set header for all worksheets
#' @description DEPRECATED
#' @author Alexander Walker
#' @param wb A workbook object
#' @param text header text. A character vector of length 1.
#' @param position Postion of text in header. One of "left", "center" or "right"
#' @export
#' @examples
#' \dontrun{
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
#' saveWorkbook(wb, "headerHeaderExample.xlsx", overwrite = TRUE)
#' }
setHeader <- function(wb, text, position = "center"){
  
  warning("This function is deprecated. Use function 'setHeaderFooter()'")
  
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
#' @description DEPRECATED
#' @author Alexander Walker
#' @param wb A workbook object
#' @param text footer text. A character vector of length 1.
#' @param position Postion of text in footer. One of "left", "center" or "right"
#' @export
#' @examples
#' \dontrun{
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
#' }
setFooter <- function(wb, text, position = "center"){
  
  warning("This function is deprecated. Use function 'setHeaderFooter()'")
  
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











#' @name dataValidation
#' @title Add data validation to cells
#' @description Add Excel data validation to cells 
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param type One of 'whole', 'decimal', 'date', 'time', 'textLength', 'list' (see examples)
#' @param operator One of 'between', 'notBetween', 'equal',
#'  'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
#' @param value a vector of length 1 or 2 depending on operator (see examples)
#' @param allowBlank logial
#' @param showInputMsg logical
#' @param showErrorMsg logical
#' @export
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#' 
#' writeDataTable(wb, 1, x = iris[1:30,])
#' 
#' dataValidation(wb, 1, col = 1:3, rows = 2:31, type = "whole"
#'    , operator = "between", value = c(1, 9))
#' 
#' dataValidation(wb, 1, col = 5, rows = 2:31, type = "textLength"
#'    , operator = "between", value = c(4, 6))
#' 
#' 
#' ## Date and Time cell validation
#' df <- data.frame("d" = as.Date("2016-01-01") + -5:5,
#'                  "t" = as.POSIXct("2016-01-01")+ -5:5*10000)
#'                  
#' writeData(wb, 2, x = df)
#' dataValidation(wb, 2, col = 1, rows = 2:12, type = "date", 
#'    operator = "greaterThanOrEqual", value = as.Date("2016-01-01"))
#'
#' dataValidation(wb, 2, col = 2, rows = 2:12, type = "time", 
#'    operator = "between", value = df$t[c(4, 8)]) 
#' 
#' saveWorkbook(wb, "dataValidationExample.xlsx", overwrite = TRUE)
#' 
#' 
#' ######################################################################
#' ## If type == 'list'
#' # operator argument is ignored.
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' addWorksheet(wb, "Sheet 2")
#' 
#' writeDataTable(wb, sheet = 1, x = iris[1:30,])
#' writeData(wb, sheet = 2, x = sample(iris$Sepal.Length, 10))
#' 
#' dataValidation(wb, 1, col = 1, rows = 2:31, type = "list", value = "'Sheet 2'!$A$1:$A$10")
#' 
#' # openXL(wb)
#' 
dataValidation <- function(wb, sheet, cols, rows, type, operator, value, allowBlank = TRUE, showInputMsg = TRUE, showErrorMsg = TRUE){
  
  ## rows and cols
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)  
  rows <- as.integer(rows)
  
  ## check length of value
  if(length(value) > 2)
    stop("value argument must be length < 2")
  
  valid_types <- c("whole", 
                   "decimal",
                   "date",
                   "time", ## need to conv
                   "textLength",
                   "list"
  )
  
  if(!tolower(type) %in% tolower(valid_types))
    stop("Invalid 'type' argument!")
  
  
  ## operator == 'between' we leave out
  valid_operators <- c("between",
                       "notBetween",
                       "equal",
                       "notEqual",
                       "greaterThan",
                       "lessThan",
                       "greaterThanOrEqual",
                       "lessThanOrEqual")
  
  if(tolower(type) != "list"){
    
    if(!tolower(operator) %in% tolower(valid_operators))
      stop("Invalid 'operator' argument!")
    
    operator <- valid_operators[tolower(valid_operators) %in% tolower(operator)][1]
    
  }else{
    operator <- "between" ## ignored
  }
  
  if(!is.logical(allowBlank))
    stop("Argument 'allowBlank' musts be logical!")
  
  if(!is.logical(showInputMsg))
    stop("Argument 'showInputMsg' musts be logical!")
  
  if(!is.logical(showErrorMsg))
    stop("Argument 'showErrorMsg' musts be logical!")
  
  ## All inputs validated
  
  type <- valid_types[tolower(valid_types) %in% tolower(type)][1]
  
  ## check input combinations
  if(type == "date" & !"Date" %in% class(value))
    stop("If type == 'date' value argument must be a Date vector.")
  
  if(type == "time" & !any(tolower(class(value)) %in% c("posixct", "posixt")))
    stop("If type == 'date' value argument must be a POSIXct or POSIXlt vector.")
  
  
  value <- head(value, 2)
  allowBlank <- as.integer(allowBlank[1])
  showInputMsg <- as.integer(showInputMsg[1])
  showErrorMsg <- as.integer(showErrorMsg[1])
  
  if(type == "list"){
    
    invisible(wb$dataValidation_list(sheet = sheet, 
                                     startRow = min(rows),
                                     endRow = max(rows),
                                     startCol = min(cols),
                                     endCol = max(cols),
                                     value = value, 
                                     allowBlank = allowBlank, 
                                     showInputMsg = showInputMsg,
                                     showErrorMsg = showErrorMsg))
    
  }else{
    
    invisible(wb$dataValidation(sheet = sheet, 
                                startRow = min(rows),
                                endRow = max(rows),
                                startCol = min(cols),
                                endCol = max(cols),
                                type = type, 
                                operator = operator, 
                                value = value, 
                                allowBlank = allowBlank, 
                                showInputMsg = showInputMsg,
                                showErrorMsg = showErrorMsg))
    
  }
  
  
  
  invisible(0)
  
}








#' @name getDateOrigin
#' @title Get the date origin an xlsx file is using
#' @description Return the date origin used internally by an xlsx or xlsm file
#' @author Alexander Walker
#' @param xlsxFile An xlsx or xlsm file.
#' @details Excel stores dates as the number of days from either 1904-01-01 or 1900-01-01. This function
#' checks the date origin being used in an Excel file and returns is so it can be used in \code{\link{convertToDate}}
#' @return One of "1900-01-01" or "1904-01-01".
#' @seealso \code{\link{convertToDate}}
#' @examples
#' 
#' ## create a file with some dates
#' write.xlsx(as.Date("2015-01-10") - (0:4), file = "getDateOriginExample.xlsx")
#' m <- read.xlsx("getDateOriginExample.xlsx")
#' 
#' ## convert to dates
#' do <- getDateOrigin(system.file("readTest.xlsx", package = "openxlsx"))
#' convertToDate(m[[1]], do)
#' 
#' @export
getDateOrigin <- function(xlsxFile){
  
  xlsxFile <- getFile(xlsxFile)
  if(!file.exists(xlsxFile))
    stop("File does not exist.")
  
  if(grepl("\\.xls$|\\.xlm$", xlsxFile))
    stop("openxlsx can not read .xls or .xlm files!")
  
  ## create temp dir and unzip
  xmlDir <- file.path(tempdir(), "_excelXMLRead")
  xmlFiles <- unzip(xlsxFile, exdir = xmlDir)
  
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)
  
  workbook <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  workbook <- paste(unlist(readLines(workbook, warn = FALSE)), collapse = "")
  
  if(grepl('date1904="1"|date1904="true"', workbook, ignore.case = TRUE)){
    origin <- "1904-01-01"
  }else{
    origin <- "1900-01-01"
  }
  
  return(origin)
  
}








#' @name getSheetNames
#' @title Get names of worksheets
#' @description Returns the worksheet names within an xlsx file
#' @author Alexander Walker
#' @param file An xlsx or xlsm file.
#' @return Character vector of worksheet names.
#' @examples
#' getSheetNames(system.file("readTest.xlsx", package = "openxlsx"))
#' 
#' @export
getSheetNames <- function(file){
  
  if(!file.exists(file))
    stop("file does not exist.")
  
  if(grepl("\\.xls$|\\.xlm$", file))
    stop("openxlsx can not read .xls or .xlm files!")
  
  ## create temp dir and unzip
  xmlDir <- file.path(tempdir(), "_excelXMLRead")
  xmlFiles <- unzip(file, exdir = xmlDir)
  
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)
  
  workbook <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  workbook <- readLines(workbook, warn=FALSE, encoding="UTF-8")
  workbook <-  removeHeadTag(workbook)
  sheets <- unlist(regmatches(workbook, gregexpr("<sheet .*/sheets>", workbook, perl = TRUE)))
  sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
  sheetNames <- replaceXMLEntities(sheetNames)
  
  return(sheetNames)
  
}




#' @name sheetVisibility
#' @title Get/set worksheet visible state
#' @description Get and set worksheet visible state
#' @param wb A workbook object 
#' @return Character vector of worksheet names.
#' @return  Vector of "hidden", "visible", "veryHidden"
#' @examples
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, sheetName = "S1", visible = FALSE)
#' addWorksheet(wb, sheetName = "S2", visible = TRUE)
#' addWorksheet(wb, sheetName = "S3", visible = FALSE)
#' 
#' sheetVisibility(wb)
#' sheetVisibility(wb)[1] <- TRUE ## show sheet 1
#' sheetVisibility(wb)[2] <- FALSE ## hide sheet 2
#' sheetVisibility(wb)[3] <- "hidden" ## hide sheet 3
#' sheetVisibility(wb)[3] <- "veryHidden" ## hide sheet 3 from UI
#' 
#' @export
sheetVisibility <- function(wb){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  state <- rep("visible", length(wb$workbook$sheets))
  state[grepl("hidden", wb$workbook$sheets)] <- "hidden"
  state[grepl("veryHidden", wb$workbook$sheets, ignore.case = TRUE)] <- "veryHidden"
  
  
  return(state)
  
}

#' @rdname sheetVisibility
#' @param value a logical/character vector the same length as sheetVisibility(wb)
#' @export
`sheetVisibility<-` <- function(wb, value) {
  
  value <- tolower(as.character(value))
  if(!any(value %in% c("true", "visible")))
    stop("A workbook must have atleast 1 visible worksheet.")
  
  value[value %in% "true"] <- "visible"
  value[value %in% "false"] <- "hidden"
  value[value %in% "veryhidden"] <- "veryHidden"
  
  
  exState0 <- regmatches(wb$workbook$sheets, regexpr('(?<=state=")[^"]+', wb$workbook$sheets, perl = TRUE))
  exState <- tolower(exState0)
  exState[exState %in% "true"] <- "visible"
  exState[exState %in% "hidden"] <- "hidden"
  exState[exState %in% "false"] <- "hidden"
  exState[exState %in% "veryhidden"] <- "veryHidden"
  
  if(length(value) != length(wb$workbook$sheets))
    stop(sprintf("value vector must have length equal to number of worksheets in Workbook [%s]", length(exState)))
  
  inds <- which(value != exState)
  if(length(inds) == 0)
    return(invisible(wb))
  
  for(i in 1:length(wb$worksheets))
    wb$workbook$sheets[i] <- gsub(exState0[i], value[i], wb$workbook$sheets[i], fixed = TRUE)
  
  invisible(wb)
  
}





#' @name pageBreak
#' @title add a page break to a worksheet
#' @description insert page breaks into a worksheet
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param i row or column number to insert page break.
#' @param type One of "row" or "column" for a row break or column break. 
#' @export
#' @seealso \code{\link{addWorksheet}}
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' writeData(wb, sheet = 1, x = iris)
#' 
#' pageBreak(wb, sheet = 1, i = 10, type = "row")
#' pageBreak(wb, sheet = 1, i = 20, type = "row")
#' pageBreak(wb, sheet = 1, i = 2, type = "column")
#' 
#' saveWorkbook(wb, "pageBreakExample.xlsx", TRUE)
#' ## In Excel: View tab -> Page Break Preview
pageBreak <- function(wb, sheet, i, type = "row"){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  sheet <- wb$validateSheet(sheet)
  
  type <- tolower(type)[1]
  if(!type %in% c("row", "column"))
    stop("'type' argument must be 'row' or 'column'.")
  
  if(!is.numeric(i))
    stop("'i' must be numeric.")
  i <- round(i)
  
  if(type == "row"){
    wb$worksheets[[sheet]]$rowBreaks <- c(
      wb$worksheets[[sheet]]$rowBreaks
      ,sprintf('<brk id="%s" max="16383" man="1"/>', i)
    )
    
  }else if(type == "column"){
    wb$worksheets[[sheet]]$colBreaks <- c(
      wb$worksheets[[sheet]]$colBreaks
      ,sprintf('<brk id="%s" max="1048575" man="1"/>', i)
    )
  }
  
  
  # wb$worksheets[[sheet]]$autoFilter <- sprintf('<autoFilter ref="%s"/>', paste(getCellRefs(data.frame("x" = c(rows, rows), "y" = c(min(cols), max(cols)))), collapse = ":"))
  
  invisible(wb)
  
}


















#' @name conditionalFormat
#' @title Add conditional formatting to cells
#' @description DEPRECATED! USE \code{\link{conditionalFormatting}}
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param rule The condition under which to apply the formatting or a vector of colours. See examples.
#' @param style A style to apply to those cells that satisify the rule. A Style object returned from createStyle()
#' @details DEPRECATED! USE \code{\link{conditionalFormatting}}
#' 
#' Valid operators are "<", "<=", ">", ">=", "==", "!=". See Examples.
#' Default style given by: createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' @param type Either 'expression', 'colorscale' or 'databar'. If 'expression' the formatting is determined
#' by a formula.  If colorScale cells are coloured based on cell value. See examples.
#' @seealso \code{\link{createStyle}}
#' @export
conditionalFormat <- function(wb, sheet, cols, rows, rule = NULL, style = NULL, type = "expression"){
  
  warning("conditionalFormat() has been deprecated. Use conditionalFormatting().")
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




#' @name all.equal
#' @aliases all.equal.Workbook
#' @title Check equality of workbooks
#' @description Check equality of workbooks
#' @method all.equal Workbook
#' @param target A \code{Workbook} object
#' @param current A \code{Workbook} object
#' @param ... ignored
all.equal.Workbook <- function(target, current, ...){
  
  
  # print("Comparing workbooks...")
  #   ".rels",
  #   "app",
  #   "charts",
  #   "colWidths",
  #   "Content_Types",
  #   "core",
  #   "drawings",
  #   "drawings_rels", 
  #   "media", 
  #   "rowHeights",
  #   "workbook",
  #   "workbook.xml.rels",
  #   "worksheets",
  #   "sheetOrder"
  #   "sharedStrings",
  #   "tables",
  #   "tables.xml.rels",
  #   "theme"
  
  
  ## TODO
  # sheet_data
  
  x <- target
  y <- current
  
  
  
  
  nSheets <- length(names(x))
  failures <- NULL
  
  flag <- all(names(x$charts) %in% names(y$charts)) & all(names(y$charts) %in% names(x$charts))
  if(!flag){
    message("charts not equal")
    failures <- c(failures, "wb$charts")
  } 
  
  flag <- all(sapply(1:nSheets, function(i) isTRUE(all.equal(x$colWidths[[i]], y$colWidths[[i]]))))
  if(!flag){
    message("colWidths not equal")
    failures <- c(failures, "wb$colWidths")
  }
  
  flag <- all(x$Content_Types %in% y$Content_Types) & all(y$Content_Types %in% x$Content_Types)
  if(!flag){
    message("Content_Types not equal")
    failures <- c(failures, "wb$Content_Types")
  } 
  
  flag <- all(unlist(x$core) == unlist(y$core))
  if(!flag){
    message("core not equal")
    failures <- c(failures, "wb$core")
  } 
  
  
  flag <- all(unlist(x$drawings) %in% unlist(y$drawings)) & all(unlist(y$drawings) %in% unlist(x$drawings))
  if(!flag){
    message("drawings not equal")
    failures <- c(failures, "wb$drawings")
  } 
  
  flag <- all(unlist(x$drawings_rels) %in% unlist(y$drawings_rels)) & all(unlist(y$drawings_rels) %in% unlist(x$drawings_rels))
  if(!flag){
    message("drawings_rels not equal")
    failures <- c(failures, "wb$drawings_rels")
  } 
  
  flag <- all(sapply(1:nSheets, function(i) isTRUE(all.equal(x$drawings_rels[[i]], y$drawings_rels[[i]]))))
  if(!flag){
    message("drawings_rels not equal")
    failures <- c(failures, "wb$drawings_rels")
  } 
  
  
  
  
  flag <- all(names(x$media) %in% names(y$media) & names(y$media) %in% names(x$media))
  if(!flag){
    message("media not equal")
    failures <- c(failures, "wb$media")
  } 
  
  flag <- all(sapply(1:nSheets, function(i) isTRUE(all.equal(x$rowHeights[[i]], y$rowHeights[[i]]))))
  if(!flag){
    message("rowHeights not equal")
    failures <- c(failures, "wb$rowHeights")
  } 
  
  flag <- all(sapply(1:nSheets, function(i) isTRUE(all.equal(names(x$rowHeights[[i]]), names(y$rowHeights[[i]])))))
  if(!flag){
    message("rowHeights not equal")
    failures <- c(failures, "wb$rowHeights")
  } 
  
  flag <- all(x$sharedStrings %in% y$sharedStrings) & all(y$sharedStrings %in% x$sharedStrings) & (length(x$sharedStrings) == length(y$sharedStrings)) 
  if(!flag){
    message("sharedStrings not equal")
    failures <- c(failures, "wb$sharedStrings")
  } 
  
  
  
  # flag <- sapply(1:nSheets, function(i) isTRUE(all.equal(x$worksheets[[i]]$sheet_data, y$worksheets[[i]]$sheet_data)))
  # if(!all(flag)){
  #   
  #   tmp_x <- x$sheet_data[[which(!flag)[[1]]]]
  #   tmp_y <- y$sheet_data[[which(!flag)[[1]]]]
  #   
  #   tmp_x_e <- sapply(tmp_x, "[[", "r")
  #   tmp_y_e <- sapply(tmp_y, "[[", "r")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("r elements: %s", paste(which(flag), collapse = ", ")))
  #     return(FALSE)
  #   }
  #   
  #   tmp_x_e <- sapply(tmp_x, "[[", "t")
  #   tmp_y_e <- sapply(tmp_y, "[[", "t")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("t elements: %s", paste(which(isTRUE(flag)), collapse = ", ")))
  #     return(FALSE)
  #   }
  #   
  #   
  #   tmp_x_e <- sapply(tmp_x, "[[", "v")
  #   tmp_y_e <- sapply(tmp_y, "[[", "v")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("v elements: %s", paste(which(flag), collapse = ", ")))
  #     return(FALSE)
  #   }
  #   
  #   tmp_x_e <- sapply(tmp_x, "[[", "f")
  #   tmp_y_e <- sapply(tmp_y, "[[", "f")
  #   flag <- paste0(tmp_x_e, "") != paste0(tmp_x_e, "")
  #   if(any(flag)){
  #     message(sprintf("sheet_data %s not equal", which(!flag)[[1]]))
  #     message(sprintf("f elements: %s", paste(which(flag), collapse = ", ")))
  #     return(FALSE)
  #   }
  # } 
  
  
  flag <- all(names(x$styles) %in% names(y$styles)) & all(names(y$styles) %in% names(x$styles))
  if(!flag){
    message("names styles not equal")
    failures <- c(failures, "names of styles not equal")
  } 
  
  flag <- all(unlist(x$styles) %in% unlist(y$styles)) & all(unlist(y$styles) %in% unlist(x$styles))
  if(!flag){
    message("styles not equal")
    failures <- c(failures, "styles not equal") 
  } 
  
  
  flag <- length(x$styleObjects) == length(y$styleObjects)
  if(!flag){
    message("styleObjects lengths not equal")
    failures <- c(failures, "styleObjects lengths not equal") 
  } 
  
  
  nStyles <- length(x$styleObjects)
  if(nStyles > 0){
    
    for(i in 1:nStyles){
      
      sx <- x$styleObjects[[i]]
      sy <- y$styleObjects[[i]]
      
      flag <- isTRUE(all.equal(sx$sheet, sy$sheet))
      if(!flag){
        message(sprintf("styleObjects '%s' sheet name not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' sheet name not equal", i)) 
      } 
      
      
      flag <- isTRUE(all.equal(sx$rows, sy$rows))
      if(!flag){
        message(sprintf("styleObjects '%s' rows not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' rows not equal", i)) 
      } 
      
      flag <- isTRUE(all.equal(sx$cols, sy$cols))
      if(!flag){
        message(sprintf("styleObjects '%s' cols not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' cols not equal", i)) 
      } 
      
      ## check style class equality
      flag <-isTRUE(all.equal(sx$style$fontName, sy$style$fontName))
      if(!flag){
        message(sprintf("styleObjects '%s' fontName not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' fontName not equal", i)) 
      } 
      
      flag <-isTRUE(all.equal(sx$style$fontColour, sy$style$fontColour))
      if(!flag){
        message(sprintf("styleObjects '%s' fontColour not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' fontColour not equal", i)) 
      }
      
      flag <-isTRUE(all.equal(sx$style$fontSize, sy$style$fontSize))
      if(!flag){
        message(sprintf("styleObjects '%s' fontSize not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' fontSize not equal", i)) 
      } 
      
      flag <-isTRUE(all.equal(sx$style$fontFamily, sy$style$fontFamily))
      if(!flag){
        message(sprintf("styleObjects '%s' fontFamily not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' fontFamily not equal", i)) 
      } 
      
      flag <-isTRUE(all.equal(sx$style$fontDecoration, sy$style$fontDecoration))
      if(!flag){
        message(sprintf("styleObjects '%s' fontDecoration not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' fontDecoration not equal", i))
      } 
      
      flag <-isTRUE(all.equal(sx$style$borderTop, sy$style$borderTop))
      if(!flag){
        message(sprintf("styleObjects '%s' borderTop not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderTop not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderLeft, sy$style$borderLeft))
      if(!flag){
        message(sprintf("styleObjects '%s' borderLeft not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderLeft not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderRight, sy$style$borderRight))
      if(!flag){
        message(sprintf("styleObjects '%s' borderRight not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderRight not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderBottom, sy$style$borderBottom))
      if(!flag){
        message(sprintf("styleObjects '%s' borderBottom not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderBottom not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderTopColour, sy$style$borderTopColour))
      if(!flag){
        message(sprintf("styleObjects '%s' borderTopColour not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderTopColour not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderLeftColour, sy$style$borderLeftColour))
      if(!flag){
        message(sprintf("styleObjects '%s' borderLeftColour not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderLeftColour not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderRightColour, sy$style$borderRightColour))
      if(!flag){
        message(sprintf("styleObjects '%s' borderRightColour not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderRightColour not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$borderBottomColour, sy$style$borderBottomColour))
      if(!flag){
        message(sprintf("styleObjects '%s' borderBottomColour not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' borderBottomColour not equal", i))
      }
      
      
      flag <-isTRUE(all.equal(sx$style$halign, sy$style$halign))
      if(!flag){
        message(sprintf("styleObjects '%s' halign not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' halign not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$valign, sy$style$valign))
      if(!flag){
        message(sprintf("styleObjects '%s' valign not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' valign not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$indent, sy$style$indent))
      if(!flag){
        message(sprintf("styleObjects '%s' indent not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' indent not equal", i))
      }
      
      
      flag <-isTRUE(all.equal(sx$style$textRotation, sy$style$textRotation))
      if(!flag){
        message(sprintf("styleObjects '%s' textRotation not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' textRotation not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$numFmt, sy$style$numFmt))
      if(!flag){
        message(sprintf("styleObjects '%s' numFmt not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' numFmt not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$fill, sy$style$fill))
      if(!flag){
        message(sprintf("styleObjects '%s' fill not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' fill not equal", i))
      }
      
      flag <-isTRUE(all.equal(sx$style$wrapText, sy$style$wrapText))
      if(!flag){
        message(sprintf("styleObjects '%s' wrapText not equal", i))
        failures <- c(failures, sprintf("styleObjects '%s' wrapText not equal", i))
      }
      
    }
    
  }
  
  
  flag <- all(x$sheet_names %in% y$sheet_names) & all(y$sheet_names %in% x$sheet_names)
  if(!flag){
    message("names workbook not equal")
    failures <- c(failures, "names workbook not equal")
  } 
  
  flag <- all(unlist(x$workbook) %in% unlist(y$workbook)) & all(unlist(y$workbook) %in% unlist(x$workbook))
  if(!flag){
    message("workbook not equal")
    failures <- c(failures, "wb$workbook")
  } 
  
  flag <- all(unlist(x$workbook.xml.rels) %in% unlist(y$workbook.xml.rels)) & all(unlist(y$workbook.xml.rels) %in% unlist(x$workbook.xml.rels))
  if(!flag){
    message("workbook.xml.rels not equal")
    failures <- c(failures, "wb$workbook.xml.rels")
  } 
  
  
  for(i in 1:nSheets){
    
    ws_x <- x$worksheets[[i]]
    ws_y <- y$worksheets[[i]]
    
    flag <- all(names(ws_x) %in% names(ws_y)) & all(names(ws_y) %in% names(ws_x))
    if(!flag){
      message(sprintf("names of worksheet elements for sheet %s not equal", i))
      failures <- c(failures, sprintf("names of worksheet elements for sheet %s not equal", i))
    } 
    
    nms <- names(ws_x)
    for(j in nms){
      
      flag <- isTRUE(all.equal(gsub(" |\t", "", ws_x[[j]]), gsub(" |\t", "", ws_y[[j]]))) 
      if(!flag){
        message(sprintf("worksheet '%s', element '%s' not equal", i, j))
        failures <- c(failures, sprintf("worksheet '%s', element '%s' not equal", i, j))
      } 
      
    }
    
  }
  
  
  flag <- all(unlist(x$sheetOrder) %in% unlist(y$sheetOrder)) & all(unlist(y$sheetOrder) %in% unlist(x$sheetOrder))
  if(!flag){
    message("sheetOrder not equal")
    failures <- c(failures, "sheetOrder not equal")
  } 
  
  
  flag <- length(x$tables) == length(y$tables)
  if(!flag){
    message("length of tables not equal")
    failures <- c(failures, "length of tables not equal")
  } 
  
  flag <- all(names(x$tables) == names(y$tables))
  if(!flag){
    message("names of tables not equal")
    failures <- c(failures, "names of tables not equal")
  } 
  
  flag <- all(unlist(x$tables) == unlist(y$tables))
  if(!flag){
    message("tables not equal")
    failures <- c(failures, "tables not equal")
  } 
  
  
  flag <- isTRUE(all.equal(x$tables.xml.rels, y$tables.xml.rels))
  if(!flag){
    message("tables.xml.rels not equal")
    failures <- c(failures, "tables.xml.rels not equal")
  } 
  
  flag <- x$theme == y$theme
  if(!flag){
    message("theme not equal")
    failures <- c(failures, "theme not equal")
  }
  
  if(!is.null(failures))
    return(FALSE)
  
  
  #   "connections",
  #   "externalLinks",
  #   "externalLinksRels",
  #   "headFoot",
  #   "pivotTables",
  #   "pivotTables.xml.rels",
  #   "pivotDefinitions",
  #   "pivotRecords",
  #   "pivotDefinitionsRels",
  #   "queryTables",
  #   "slicers",
  #   "slicerCaches",
  #   "vbaProject",
  
  
  return(TRUE)
}



#' @name sheetVisible
#' @title Get worksheet visible state.
#' @description DEPRECATED - Use function 'sheetVisibility()
#' @author Alexander Walker
#' @param wb A workbook object 
#' @return Character vector of worksheet names.
#' @return  TRUE if sheet is visible, FALSE if sheet is hidden
#' @examples
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, sheetName = "S1", visible = FALSE)
#' addWorksheet(wb, sheetName = "S2", visible = TRUE)
#' addWorksheet(wb, sheetName = "S3", visible = FALSE)
#' 
#' sheetVisible(wb)
#' sheetVisible(wb)[1] <- TRUE ## show sheet 1
#' sheetVisible(wb)[2] <- FALSE ## hide sheet 2
#' 
#' @export
sheetVisible <- function(wb){
  
  warning("This function is deprecated. Use function 'sheetVisibility()'")
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  state <- rep(TRUE, length(wb$workbook$sheets))
  state[grepl("hidden", wb$workbook$sheets)] <- FALSE
  
  return(state)
  
}

#' @rdname sheetVisible
#' @param value a logical vector the same length as sheetVisible(wb)
#' @export
`sheetVisible<-` <- function(wb, value) {
  
  warning("This function is deprecated. Use function 'sheetVisibility()'")
  
  if(!is.logical(value))
    stop("value must be a logical vector.")
  
  if(!any(value))
    stop("A workbook must have atleast 1 visible worksheet.")
  
  value <- as.character(value)
  value[value %in% "TRUE"] <- "visible"
  value[value %in% "FALSE"] <- "hidden"
  
  exState <- rep("visible", length(wb$workbook$sheets))
  exState[grepl("hidden", wb$workbook$sheets)] <- "hidden"
  
  if(length(value) != length(wb$workbook$sheets))
    stop(sprintf("value vector must have length equal to number of worksheets in Workbook [%s]", length(exState)))
  
  inds <- which(value != exState)
  if(length(inds) == 0)
    return(invisible(wb))
  
  for(i in inds)
    wb$workbook$sheets[i] <- gsub(exState[i], value[i], wb$workbook$sheets[i])
  
  invisible(wb)
  
}



#' @name copyWorkbook
#' @title Copy a Workbook object.
#' @description Just a wrapper of wb$copy() 
#' @param wb A workbook object 
#' @return Workbook
#' @examples
#' 
#' wb <- createWorkbook()  
#' wb2 <- wb ## does not create a copy
#' wb3 <- copyWorkbook(wb) ## wrapper for wb$copy()
#' 
#' addWorksheet(wb, "Sheet1") ## adds worksheet to both wb and wb2 but not wb3
#' 
#' names(wb)
#' names(wb2)
#' names(wb3)
#' 
#' @export
copyWorkbook <- function(wb){
  
  if(!inherits(wb, "Workbook"))
    stop("argument must be a Workbook.")
  
  return(wb$copy())
  
}





#' @name getTables
#' @title List Excel tables in a workbook
#' @description List Excel tables in a workbook
#' @param wb A workbook object 
#' @param sheet A name or index of a worksheet
#' @return character vector of table names on the specified sheet
#' @examples
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, sheetName = "Sheet 1")
#' writeDataTable(wb, sheet = "Sheet 1", x = iris)
#' writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#' 
#' getTables(wb, sheet = "Sheet 1")
#' 
#' 
#' @export
getTables <- function(wb, sheet){
  
  if(!inherits(wb, "Workbook"))
    stop("argument must be a Workbook.")
  
  if(length(sheet) != 1)
    stop("sheet argument must be length 1")
  
  if(length(wb$tables) == 0)
    return(character(0))
  
  sheet <- wb$validateSheet(sheetName = sheet)
  
  table_sheets <- attr(wb$tables, "sheet")
  tables <- attr(wb$tables, "tableName")
  refs <- names(wb$tables)
  
  refs <- refs[table_sheets == sheet & !grepl("openxlsx_deleted", tables, fixed = TRUE)]
  tables <- tables[table_sheets == sheet & !grepl("openxlsx_deleted", tables, fixed = TRUE)]
  
  if(length(tables) > 0)
    attr(tables, "refs") <- refs
  
  return(tables)
  
}





#' @name removeTable
#' @title Remove an Excel table in a workbook
#' @description List Excel tables in a workbook
#' @param wb A workbook object 
#' @param sheet A name or index of a worksheet
#' @param table Name of table to remove. See \code{\link{getTables}}
#' @return character vector of table names on the specified sheet
#' @examples
#' 
#' wb <- createWorkbook()
#' addWorksheet(wb, sheetName = "Sheet 1")
#' addWorksheet(wb, sheetName = "Sheet 2")
#' writeDataTable(wb, sheet = "Sheet 1", x = iris, tableName = "iris")
#' writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#' 
#' 
#' removeWorksheet(wb, sheet = 1) ## delete worksheet removes table objects
#' 
#' writeDataTable(wb, sheet = 1, x = iris, tableName = "iris")
#' writeDataTable(wb, sheet = 1, x = mtcars, tableName = "mtcars", startCol = 10)
#' 
#' ## removeTable() deletes table object and all data
#' getTables(wb, sheet = 1)
#' removeTable(wb = wb, sheet = 1, table = "iris")
#' writeDataTable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
#' 
#' getTables(wb, sheet = 1)
#' removeTable(wb = wb, sheet = 1, table = "iris")
#' writeDataTable(wb, sheet = 1, x = iris, tableName = "iris", startCol = 1)
#' 
#' saveWorkbook(wb = wb, file = "removeTableExample.xlsx", overwrite = TRUE)
#'  
#' @export
removeTable <- function(wb, sheet, table){
  
  if(!inherits(wb, "Workbook"))
    stop("argument must be a Workbook.")
  
  if(length(sheet) != 1)
    stop("sheet argument must be length 1")
  
  if(length(table) != 1)
    stop("table argument must be length 1")
  
  ## delete table object and all data in it
  sheet <- wb$validateSheet(sheetName = sheet)

  if(!table %in% attr(wb$tables, "tableName"))
    stop(sprintf("table '%s' does not exist.", table), call.=FALSE)
  
  ## get existing tables
  table_sheets <- attr(wb$tables, "sheet")
  table_names <- attr(wb$tables, "tableName")
  refs <- names(wb$tables)
  
  ## delete table object (by flagging as deleted)
  inds <- which(table_sheets %in% sheet & table_names %in% table)
  table_name_original <- table_names[inds]
  
  table_names[inds] <- paste0(table_name_original, "_openxlsx_deleted")
  attr(wb$tables, "tableName") <- table_names
  
  ## delete reference from worksheet to table
  worksheet_table_names <- attr(wb$worksheets[[sheet]]$tableParts, "tableName")
  to_remove <- which(worksheet_table_names == table_name_original)
  
  wb$worksheets[[sheet]]$tableParts <- wb$worksheets[[sheet]]$tableParts[-to_remove]
  attr(wb$worksheets[[sheet]]$tableParts, "tableName") <- worksheet_table_names[-to_remove]
  
  
  ## Now delete data from the worksheet
  refs <- strsplit(refs[[inds]], split = ":")[[1]]
  rows <- as.integer(gsub("[A-Z]", "", refs))
  rows <- seq(from = rows[1], to = rows[2], by = 1)
  
  cols <- convertFromExcelRef(refs)
  cols <- seq(from = cols[1], to = cols[2], by = 1)
  
  ## now delete data
  deleteData(wb = wb, sheet = sheet, rows = rows, cols = cols, gridExpand = TRUE)
  
  invisible(0)
  
  
}
