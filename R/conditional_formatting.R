



#' @name conditionalFormatting
#' @aliases databar
#' @title Add conditional formatting to cells
#' @description Add conditional formatting to cells
#' @author Alexander Walker
#' @param wb A workbook object
#' @param sheet A name or index of a worksheet
#' @param cols Columns to apply conditional formatting to
#' @param rows Rows to apply conditional formatting to
#' @param rule The condition under which to apply the formatting. See examples.
#' @param style A style to apply to those cells that satisify the rule. Default is createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' @param type Either 'expression', 'colorscale', 'databar', 'duplicates' or "contains' (case insensitive).
#' @param ... See below 
#' @details See Examples.
#' 
#' If type == "expression"
#' \itemize{
#'   \item{style is a Style object. See \code{\link{createStyle}}}
#'   \item{rule is an expression. Valid operators are "<", "<=", ">", ">=", "==", "!=".}
#' }
#' 
#' If type == "colourScale"
#' \itemize{
#'   \item{style is a vector of colours with length 2 or 3}
#'   \item{rule can be NULL or a vector of colours of equal length to styles}
#' }
#' 
#' If type == "databar"
#' \itemize{
#'   \item{style is a vector of colours with length 2 or 3}
#'   \item{rule is a numeric vector specifiying the range of the databar colours. Must be equal length to style}
#'   \item{...
#'   \itemize{
#'     \item{\bold{showvalue} If FALSE the cell value is hidden. Default TRUE.}
#'     \item{\bold{gradient} If FALSE colour gradient is removed. Default TRUE.}
#'     \item{\bold{border} If FALSE the border around the database is hidden. Default TRUE.}
#'      }
#'    }
#' }
#' 
#' If type == "duplicates"
#' \itemize{
#'   \item{style is a Style object. See \code{\link{createStyle}}}
#'   \item{rule is ignored.}
#' }
#' 
#' If type == "contains"
#' \itemize{
#'   \item{style is a Style object. See \code{\link{createStyle}}}
#'   \item{rule is the text to look for within cells}
#' }
#' 
#' If type == "between"
#' \itemize{
#'   \item{style is a Style object. See \code{\link{createStyle}}}
#'   \item{rule is a numeric vector of length 2 specifying lower and upper bound (Inclusive)}
#' }
#' 
#' @seealso \code{\link{createStyle}}
#' @export
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "cellIs")
#' addWorksheet(wb, "Moving Row")
#' addWorksheet(wb, "Moving Col")
#' addWorksheet(wb, "Dependent on 1")
#' addWorksheet(wb, "Duplicates")
#' addWorksheet(wb, "containsText")
#' addWorksheet(wb, "colourScale", zoom = 30)
#' addWorksheet(wb, "databar")
#' addWorksheet(wb, "between")
#' addWorksheet(wb, "logical operators")
#' 
#' negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
#' posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
#' 
#' ## rule applies to all each cell in range
#' writeData(wb, "cellIs", -5:5)
#' writeData(wb, "cellIs", LETTERS[1:11], startCol=2)
#' conditionalFormatting(wb, "cellIs", cols=1, rows=1:11, rule="!=0", style = negStyle)
#' conditionalFormatting(wb, "cellIs", cols=1, rows=1:11, rule="==0", style = posStyle)
#' 
#' ## highlight row dependent on first cell in row
#' writeData(wb, "Moving Row", -5:5)
#' writeData(wb, "Moving Row", LETTERS[1:11], startCol=2)
#' conditionalFormatting(wb, "Moving Row", cols=1:2, rows=1:11, rule="$A1<0", style = negStyle)
#' conditionalFormatting(wb, "Moving Row", cols=1:2, rows=1:11, rule="$A1>0", style = posStyle)
#' 
#' ## highlight column dependent on first cell in column
#' writeData(wb, "Moving Col", -5:5)
#' writeData(wb, "Moving Col", LETTERS[1:11], startCol=2)
#' conditionalFormatting(wb, "Moving Col", cols=1:2, rows=1:11, rule="A$1<0", style = negStyle)
#' conditionalFormatting(wb, "Moving Col", cols=1:2, rows=1:11, rule="A$1>0", style = posStyle)
#' 
#' ## highlight entire range cols X rows dependent only on cell A1
#' writeData(wb, "Dependent on 1", -5:5)
#' writeData(wb, "Dependent on 1", LETTERS[1:11], startCol=2)
#' conditionalFormatting(wb, "Dependent on 1", cols=1:2, rows=1:11, rule="$A$1<0", style = negStyle)
#' conditionalFormatting(wb, "Dependent on 1", cols=1:2, rows=1:11, rule="$A$1>0", style = posStyle)
#' 
#' ## highlight duplicates using default style
#' writeData(wb, "Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
#' conditionalFormatting(wb, "Duplicates", cols = 1, rows = 1:10, type = "duplicates")
#' 
#' ## cells containing text
#' fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
#' writeData(wb, "containsText", sapply(1:10, fn))
#' conditionalFormatting(wb, "containsText", cols = 1, rows = 1:10, type = "contains", rule = "A")
#' 
#' ## colourscale colours cells based on cell value
#' df <- read.xlsx(system.file("readTest.xlsx", package = "openxlsx"), sheet = 4)
#' writeData(wb, "colourScale", df, colNames=FALSE)  ## write data.frame
#' 
#' ## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
#' ## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
#' conditionalFormatting(wb, "colourScale", cols=1:ncol(df), rows=1:nrow(df),
#'    style = c("black", "white"), 
#'    rule = c(0, 255), 
#'    type = "colourScale")
#' 
#' setColWidths(wb, "colourScale", cols = 1:ncol(df), widths = 1.07)
#' setRowHeights(wb, "colourScale", rows = 1:nrow(df), heights = 7.5) 
#' 
#' ## Databars
#' writeData(wb, "databar", -5:5)
#' conditionalFormatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Default colours
#' 
#' ## Betweem
#' # Highlight cells in interval [-2, 2]
#' writeData(wb, "between", -5:5)
#' conditionalFormatting(wb, "between", cols = 1, rows = 1:11, type = "between", rule = c(-2,2))
#' 
#' ## Logical Operators
#' # You can use Excels logical Opertors
#' writeData(wb, "logical operators", 1:10)
#' conditionalFormatting(wb, "logical operators", cols = 1, rows = 1:10,
#'  rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)")
#' 
#' saveWorkbook(wb, "conditionalFormattingExample.xlsx", TRUE)
#' 
#' 
#' #########################################################################
#' ## Databar Example
#' 
#' wb <- createWorkbook()
#'addWorksheet(wb, "databar")
#'
#'## Databars
#'writeData(wb, "databar", -5:5, startCol = 1)
#'conditionalFormatting(wb, "databar", cols = 1, rows = 1:11, type = "databar") ## Defaults
#' 
#' writeData(wb, "databar", -5:5, startCol = 3)
#' conditionalFormatting(wb, "databar", cols = 3, rows = 1:11, type = "databar", border = FALSE)
#' 
#' writeData(wb, "databar", -5:5, startCol = 5)
#' conditionalFormatting(wb, "databar", cols = 5, rows = 1:11, 
#'   type = "databar", style = c("#a6a6a6"), showValue = FALSE) 
#' 
#' writeData(wb, "databar", -5:5, startCol = 7)
#' conditionalFormatting(wb, "databar", cols = 7, rows = 1:11, 
#'   type = "databar", style = c("#a6a6a6"), showValue = FALSE, gradient = FALSE) 
#' 
#' writeData(wb, "databar", -5:5, startCol = 9)
#' conditionalFormatting(wb, "databar", cols = 9, rows = 1:11, 
#'   type = "databar", style = c("#a6a6a6", "#a6a6a6"), showValue = FALSE, gradient = FALSE)
#' 
#' saveWorkbook(wb, file = "databarExample.xlsx", overwrite = TRUE)
#'  
#' 
conditionalFormatting <- function(wb, sheet, cols, rows, rule = NULL, style = NULL, type = "expression", ...){
  
  type <- tolower(type)
  params <- list(...)
  
  if(type %in% c("colorscale", "colourscale")){
    type <- "colorScale"
    
  }else if(type == "databar"){
    type <- "dataBar"
    
  }else if(type == "duplicates"){
    type <- "duplicatedValues"
    
  }else if(type == "contains"){
    type <- "containsText"
    
  }else if(type == "between"){
    type <- "between"
    
  }else if(type != "expression"){
    stop("Invalid type argument.  Type must be one of 'expression', 'colourScale', 'databar', 'duplicates' or 'contains'")
  }
  
  ## rows and cols
  if(!is.numeric(cols))
    cols <- convertFromExcelRef(cols)  
  rows <- as.integer(rows)
  
  
  ## check valid rule
  values <- NULL
  dxfId <- NULL
  
  if(type == "colorScale"){
    
    # type == "colourScale"
    # - style is a vector of colours with length 2 or 3
    # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used
    
    if(is.null(style))
      stop("If type == 'colourScale', style must be a vector of colours of length 2 or 3.")
    
    if(class(style) != "character")
      stop("If type == 'colourScale', style must be a vector of colours of length 2 or 3.")
    
    if(!length(style) %in% 2:3)
      stop("If type == 'colourScale', style must be a vector of length 2 or 3.")
    
    if(!is.null(rule)){
      if(length(rule) != length(style))
        stop("If type == 'colourScale', rule and style must have equal lengths.")
    }
    
    style <- validateColour(style, errorMsg="Invalid colour specified in style.")
    
    values <- rule
    rule <- style
    
  }else if(type == "dataBar"){
    
    # type == "databar"
    # - style is a vector of colours of length 2 or 3
    # - rule specifies the quantiles (numeric vector of length 2 or 3), if NULL min and max are used
    
    if(is.null(style))
      style <- "#638EC6"
    
    if(class(style) != "character")
      stop("If type == 'dataBar', style must be a vector of colours of length 1 or 2.")
    
    if(!length(style) %in% 1:2)
      stop("If type == 'dataBar', style must be a vector of length 1 or 2.")
    
    if(!is.null(rule)){
      if(length(rule) != length(style))
        stop("If type == 'dataBar', rule and style must have equal lengths.")
    }
    
    
    ## Additional paramters passed by ...
    if("showValue" %in% names(params)){
      params$showValue <- as.integer(params$showValue)
      if(is.na(params$showValue))
        stop("showValue must be 0/1 or TRUE/FALSE")
    }
    
    if("gradient" %in% names(params)){
      params$gradient <- as.integer(params$gradient)
      if(is.na(params$gradient))
        stop("gradient must be 0/1 or TRUE/FALSE")
    }
    
    if("border" %in% names(params)){
      params$border <- as.integer(params$border)
      if(is.na(params$border))
        stop("border must be 0/1 or TRUE/FALSE")
    }

    style <- validateColour(style, errorMsg="Invalid colour specified in style.")
    
    values <- rule
    rule <- style
    
  }else if(type == "expression"){
    
    # type == "expression"
    # - style = createStyle()
    # - rule is an expression to evaluate
    
    # rule <- gsub(" ", "", rule)
    rule <- replaceIllegalCharacters(rule)
    rule <- gsub("!=", "&lt;&gt;", rule)
    rule <- gsub("==", "=", rule)
    
    if(!grepl("[A-Z]", substr(rule, 1, 2))){
      
      ## formula looks like "operatorX" , attach top left cell to rule    
      rule <- paste0( getCellRefs(data.frame("x" = min(rows), "y" = min(cols))), rule)
      
    } ## else, there is a letter in the formula and apply as is
    
    if(is.null(style))
      style <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
    
    if(!"Style" %in% class(style))
      stop("If type == 'expression', style must be a Style object.")
    
    invisible(dxfId <- wb$addDXFS(style))
    
  }else if(type == "duplicatedValues"){
    
    # type == "duplicatedValues"
    # - style is a Style object
    # - rule is ignored
    
    if(is.null(style))
      style <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
    
    if(!"Style" %in% class(style))
      stop("If type == 'duplicates', style must be a Style object.")
    
    invisible(dxfId <- wb$addDXFS(style))
    rule <- style
    
  }else if(type == "containsText"){
    
    # type == "contains"
    # - style is Style object
    # - rule is text to look for
    
    if(is.null(style))
      style <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
    
    if(!"character" %in% class(rule))
      stop("If type == 'contains', rule must be a character vector of length 1.")
    
    if(!"Style" %in% class(style))
      stop("If type == 'contains', style must be a Style object.")
    
    invisible(dxfId <- wb$addDXFS(style))
    values <- rule
    rule <- style
    
  }else if(type == "between"){
    
    rule <- range(rule)
    
    if(is.null(style))
      style <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
    
    if(!"Style" %in% class(style))
      stop("If type == 'between', style must be a Style object.")
    
    invisible(dxfId <- wb$addDXFS(style))
  }
  
  
  
  invisible(wb$conditionalFormatting(sheet,
                                     startRow = min(rows),
                                     endRow = max(rows),
                                     startCol = min(cols),
                                     endCol = max(cols),
                                     dxfId = dxfId,
                                     formula = rule,
                                     type = type,
                                     values = values,
                                     params = params))
  
  invisible(0)
  
}


