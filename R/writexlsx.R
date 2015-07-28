

#' @name write.xlsx
#' @title write directly to an xlsx file
#' @author Alexander Walker
#' @param x object or a list of objects that can be handled by \code{\link{writeData}} to write to file
#' @param file xlsx file name
#' @param asTable write using writeDataTable as opposed to writeData
#' @param ... optional parameters to pass to functions:
#' \itemize{
#'   \item{createWorkbook}
#'   \item{addWorksheet}
#'   \item{writeData}
#'   \item{freezePane}
#'   \item{saveWorkbook}
#' }
#' 
#' see details.
#' @details Optional parameters are:
#'
#' \bold{createWorkbook Parameters}
#' \itemize{
#'   \item{\bold{creator}}{ A string specifying the workbook author}
#' }
#' 
#' \bold{addWorksheet Parameters}
#' \itemize{
#'   \item{\bold{sheetName}}{ Name of the worksheet}
#'   \item{\bold{gridLines}}{ A logical. If \code{FALSE}, the worksheet grid lines will be hidden.}
#'   \item{\bold{tabColour}}{ Colour of the worksheet tab. A valid colour (belonging to colours()) 
#'   or a valid hex colour beginning with "#".}
#'   \item{\bold{zoom}}{ A numeric betwettn 10 and 400. Worksheet zoom level as a percentage.}
#' }
#' 
#' \bold{writeData/writeDataTable Parameters}
#' \itemize{
#'   \item{\bold{startCol}}{ A vector specifiying the starting column(s) to write df}
#'   \item{\bold{startRow}}{ A vector specifiying the starting row(s) to write df}
#'   \item{\bold{xy}}{ An alternative to specifying startCol and startRow individually. 
#'  A vector of the form c(startCol, startRow)}
#'   \item{\bold{colNames or col.names}}{ If \code{TRUE}, column names of x are written.}
#'   \item{\bold{rowNames or row.names}}{ If \code{TRUE}, row names of x are written.}
#'   \item{\bold{headerStyle}}{ Custom style to apply to column names.}
#'   \item{\bold{borders}}{ Either "surrounding", "columns" or "rows" or NULL.  If "surrounding", a border is drawn around the
#' data.  If "rows", a surrounding border is drawn a border around each row. If "columns", a surrounding border is drawn with a border
#' between each column.  If "\code{all}" all cell borders are drawn.}
#'   \item{\bold{borderColour}}{ Colour of cell border}
#'   \item{\bold{borderStyle}}{ Border line style.}
#'   \item{\bold{keepNA}} {If \code{TRUE}, NA values are converted to #N/A in Excel else NA cells will be empty. Defaults to FALSE.}
#' }
#' 
#' \bold{freezePane Parameters}
#' \itemize{
#'   \item{\bold{firstActiveRow}} {Top row of active region to freeze pane.}
#'   \item{\bold{firstActiveCol}} {Furthest left column of active region to freeze pane.}
#'   \item{\bold{firstRow}} {If \code{TRUE}, freezes the first row (equivalent to firstActiveRow = 2)}
#'   \item{\bold{firstCol}} {If \code{TRUE}, freezes the first column (equivalent to firstActiveCol = 2)}
#' }
#' 
#' \bold{colWidths Parameters}
#' \itemize{
#'   \item{\bold{colWidths}} {Must be value "auto". Sets all columns containing data to auto width.}
#' }
#' 
#' 
#' \bold{saveWorkbook Parameters}
#' \itemize{
#'   \item{\bold{overwrite}}{ Overwrite existing file (Defaults to TRUE as with write.table)}
#' }
#' 
#' 
#' columns of x with class Date or POSIXt are automatically
#' styled as dates and datetimes respectively.
#' @seealso \code{\link{addWorksheet}}
#' @seealso \code{\link{writeData}}
#' @seealso \code{\link{createStyle}} for style parameters
#' @return A workbook object
#' @examples
#' 
#' ## write to working directory
#' options("openxlsx.borderColour" = "#4F80BD") ## set default border colour
#' write.xlsx(iris, file = "writeXLSX1.xlsx", colNames = TRUE, borders = "columns")
#' write.xlsx(iris, file = "writeXLSX2.xlsx", colNames = TRUE, borders = "surrounding")
#' 
#' 
#' hs <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,
#'                   fontName="Arial Narrow", fgFill = "#4F80BD")
#' 
#' write.xlsx(iris, file = "writeXLSX3.xlsx", colNames = TRUE, borders = "rows", headerStyle = hs)
#' 
#' ## Lists elements are written to individual worksheets, using list names as sheet names if available
#' l <- list("IRIS" = iris, "MTCATS" = mtcars, matrix(runif(1000), ncol = 5))
#' write.xlsx(l, "writeList1.xlsx", colWidths = c(NA, "auto", "auto"))
#' 
#' ## different sheets can be given different parameters
#' write.xlsx(l, "writeList2.xlsx", startCol = c(1,2,3), startRow = 2,
#'            asTable = c(TRUE, TRUE, FALSE), withFilter = c(TRUE, FALSE, FALSE))
#' 
#' @export
write.xlsx <- function(x, file, asTable = FALSE, ...){
  
  
  ## set scientific notation penalty
  
  params <- list(...)
  
  ## Possible parameters
  
  #---createWorkbook---#
  ## creator
  
  #---addWorksheet---#
  ## sheetName
  ## gridLines
  ## tabColour = NULL
  ## zoom = 100
  ## header = NULL
  ## footer = NULL
  ## evenHeader = NULL
  ## evenFooter = NULL
  ## firstHeader = NULL
  ## firstFooter = NULL
  
  #---writeData---#
  ## startCol = 1,
  ## startRow = 1, 
  ## xy = NULL,
  ## colNames = TRUE,
  ## rowNames = FALSE,
  ## headerStyle = NULL,
  ## borders = NULL,
  ## borderColour = "#4F81BD"
  ## borderStyle
  ## keepNA = FALSE
  
  #----writeDataTable---#
  ## startCol = 1
  ## startRow = 1
  ## xy = NULL
  ## colNames = TRUE
  ## rowNames = FALSE
  ## tableStyle = "TableStyleLight9"
  ## tableName = NULL
  ## headerStyle = NULL
  ## withFilter = TRUE
  
  #---freezePane---#
  ## firstActiveRow = NULL
  ## firstActiveCol = NULL
  ## firstRow = FALSE
  ## firstCol = FALSE
  
  
  #---saveWorkbook---#
  #   overwrite = TRUE
  
  if(!is.logical(asTable))
    stop("asTable must be a logical.")
  
  creator <- ""
  if(creator %in% names(params))
    creator <- params$creator
  
  sheetName <- "Sheet 1"
  if("sheetName" %in% names(params)){
    
    if(nchar(params$sheetName) > 31)
      stop("sheetName too long! Max length is 31 characters.")
    
    sheetName <- as.character(params$sheetName)
  }
  
  tabColour <- NULL
  if("tabColour" %in% names(params))
    tabColour <- validateColour(params$tabColour, "Invalid tabColour!")
  
  zoom <- 100
  if("zoom" %in% names(params)){
    if(is.numeric(params$zoom)){
      zoom <- params$zoom[1]
    }else{
      stop("zoom must be numeric")
    }
  }
  
  ## AddWorksheet
  gridLines <- TRUE
  if("gridLines" %in% names(params)){
    if(all(is.logical(params$gridLines))){
      gridLines <- params$gridLines
    }else{
      stop("Argument gridLines must be TRUE or FALSE")
    }
  }
  
  overwrite <- TRUE
  if("overwrite" %in% names(params)){
    if(is.logical(params$overwrite)){
      overwrite <- params$overwrite
    }else{
      stop("Argument overwrite must be TRUE or FALSE")
    }
  }
  
  
  withFilter <- TRUE
  if("withFilter" %in% names(params)){
    if(is.logical(params$withFilter)){
      withFilter <- params$withFilter
    }else{
      stop("Argument withFilter must be TRUE or FALSE")
    }
  }
  
  startRow <- 1
  if("startRow" %in% names(params)){
    if(all(startRow > 0)){
      startRow <- params$startRow
    }else{
      stop("startRow must be a positive integer")
    }
  }
  
  startCol <- 1
  if("startCol" %in% names(params)){
    if(all(startCol > 0)){
      startCol <- params$startCol
    }else{
      stop("startCol must be a positive integer")
    }
  }
  
  colNames <- TRUE
  if("colNames" %in% names(params)){
    if(is.logical(params$colNames)){
      colNames <- params$colNames
    }else{
      stop("Argument colNames must be TRUE or FALSE")
    }
  }
  
  ## to be consistent with write.csv
  if("col.names" %in% names(params)){
    if(is.logical(params$col.names)){
      colNames <- params$col.names
    }else{
      stop("Argument col.names must be TRUE or FALSE")
    }
  }
  
  
  rowNames <- FALSE
  if("rowNames" %in% names(params)){
    if(is.logical(params$rowNames)){
      rowNames <- params$rowNames
    }else{
      stop("Argument colNames must be TRUE or FALSE")
    }
  }
  
  ## to be consistent with write.csv
  if("row.names" %in% names(params)){
    if(is.logical(params$row.names)){
      rowNames <- params$row.names
    }else{
      stop("Argument row.names must be TRUE or FALSE")
    }
  }
  
  xy <- NULL
  if("xy" %in% names(params)){
    if(length(params$xy) != 2)
      stop("xy parameter must have length 2")
    xy <- params$xy
  }
  
  
  headerStyle <- NULL
  if("headerStyle" %in% names(params)){
    
    if(length(params$headerStyle) == 1){
      if("Style" %in% class(params$headerStyle)){
        headerStyle <- params$headerStyle
      }else{
        stop("headerStyle must be a style object.")
      }
    }else{
      if(all(sapply(params$headerStyle, function(x) "Style" %in% class(x)))){
        headerStyle <- params$headerStyle
      }else{
        stop("headerStyle must be a style object.")
      }
    }
  }
  
  borders <- NULL
  if("borders" %in% names(params)){
    borders <- tolower(params$borders)
    if(!all(borders %in% c("surrounding", "rows", "columns")))
      stop("Invalid borders argument")
  }
  
  borderColour <- getOption("openxlsx.borderColour", "black")
  if("borderColour" %in% names(params))
    borderColour <- params$borderColour
  
  borderStyle <- getOption("openxlsx.borderStyle", "thin")
  if("borderStyle" %in% names(params)){
    borderStyle <- validateBorderStyle(params$borderStyle)
  }
  
  keepNA <- FALSE
  if("keepNA" %in% names(params)){
    if(!"logical" %in% class(keepNA)){
      stop("keepNA must be a logical.")
    }else{
      keepNA <- params$keepNA
    }
  }
  
  
  tableStyle <- "TableStyleLight9"
  if("tableStyle" %in% names(params))
    tableStyle <- params$tableStyle
  
  
  ## auto column widths
  colWidths <- ""
  if("colWidths" %in% names(params))
    colWidths <- params$colWidths
  
    
  ## create new Workbook object
  wb <- Workbook$new(creator)
  
  ## If a list is supplied write to individual worksheets using names if available
  nSheets <- 1
  if("list" %in% class(x)){
    
    nms <- names(x)
    nSheets <- length(x)
    
    if(is.null(nms)){
      nms <- paste("Sheet", 1:nSheets)
    }else if(any("" %in% nms)){
      nms[nms %in% ""] <- paste("Sheet", (1:nSheets)[nms %in% ""])
    }else{
      nms <- make.names(nms, unique  = TRUE)
    }
    
    if(any(nchar(nms) > 31)){
      warning("Truncating list names to 31 characters.")
      nms <- substr(nms, 1, 31)
    }
    
    ## make all inputs as long as the list
    if(length(withFilter) != nSheets)
      withFilter <- rep_len(withFilter, length.out = nSheets)
    
    if(length(colNames) != nSheets)
      colNames <- rep_len(colNames, length.out = nSheets)
    
    if(length(rowNames) != nSheets)
      rowNames <- rep_len(rowNames, length.out = nSheets)
    
    if(length(startRow) != nSheets)
      startRow <- rep_len(startRow, length.out = nSheets)
    
    if(length(startCol) != nSheets)
      startCol <- rep_len(startCol, length.out = nSheets)
    
    if(length(headerStyle) != nSheets & !is.null(headerStyle))
      headerStyle <- lapply(1:nSheets, function(x) headerStyle)
    
    if(length(borders) != nSheets & !is.null(borders))
      borders <- rep_len(borders, length.out = nSheets)
    
    if(length(borderColour) != nSheets)
      borderColour <- rep_len(borderColour, length.out = nSheets)
    
    if(length(borderStyle) != nSheets)
      borderStyle <- rep_len(borderStyle, length.out = nSheets)
    
    if(length(keepNA) != nSheets)
      keepNA <- rep_len(keepNA, length.out = nSheets)
    
    if(length(asTable) != nSheets)
      asTable <- rep_len(asTable, length.out = nSheets)
    
    if(length(tableStyle) != nSheets)
      tableStyle <- rep_len(tableStyle, length.out = nSheets)
    
    if(length(colWidths) != nSheets)
      colWidths <- rep_len(colWidths, length.out = nSheets)
    
    for(i in 1:nSheets){
      
      wb$addWorksheet(nms[[i]], showGridLines = gridLines, tabColour = tabColour, zoom = zoom)
      
      if(asTable[i]){
        
        writeDataTable(wb = wb,
                       sheet = i,
                       x = x[[i]],
                       startCol = startCol[[i]],
                       startRow = startRow[[i]],
                       xy = xy,
                       colNames = colNames[[i]],
                       rowNames = rowNames[[i]],
                       tableStyle = tableStyle[[i]],
                       tableName = NULL,
                       headerStyle = headerStyle[[i]],
                       withFilter = withFilter[[i]],
                       keepNA = keepNA[[i]])
        
      }else{
        
        writeData(wb = wb, 
                  sheet = i,
                  x = x[[i]],
                  startCol = startCol[[i]],
                  startRow = startRow[[i]], 
                  xy = xy,
                  colNames = colNames[[i]],
                  rowNames = rowNames[[i]],
                  headerStyle = headerStyle[[i]],
                  borders = borders[[i]],
                  borderColour = borderColour[[i]],
                  borderStyle = borderStyle[[i]],
                  keepNA = keepNA[[i]])
        
      }
      
      if(colWidths[i] %in% "auto")
        setColWidths(wb, sheet = i, cols =  1:ncol(x[[i]]) +  startCol[[i]] - 1L, widths = "auto")
      
      
      
    }
    
    
  }else{
    
    wb$addWorksheet(sheetName, showGridLines = gridLines, tabColour = tabColour, zoom = zoom)
    
    if(asTable){
      
      if(!"data.frame" %in% class(x))
        stop("x must be a data.frame is asTable == TRUE")
      
      writeDataTable(wb = wb,
                     sheet = 1,
                     x = x,
                     startCol = startCol,
                     startRow = startRow,
                     xy = xy,
                     colNames = colNames,
                     rowNames = rowNames,
                     tableStyle = tableStyle,
                     tableName = NULL,
                     headerStyle = headerStyle,
                     keepNA = keepNA)
      
    }else{
      
      writeData(wb = wb, 
                sheet = 1,
                x = x,
                startCol = startCol,
                startRow = startRow, 
                xy = xy,
                colNames = colNames,
                rowNames = rowNames,
                headerStyle = headerStyle,
                borders = borders,
                borderColour = borderColour,
                borderStyle = borderStyle,
                keepNA = keepNA)      
    }
    
    if(colWidths[1] %in% "auto")
      setColWidths(wb, sheet = 1, cols =  1:ncol(x) +  startCol - 1L, widths = "auto")
    
  }
  
  ###--Freeze Panes---###
  ## firstActiveRow = NULL
  ## firstActiveCol = NULL
  ## firstRow = FALSE
  ## firstCol = FALSE
  
  freezePanes <- FALSE
  firstActiveRow <- rep_len(1L, length.out = nSheets)
  if("firstActiveRow" %in% names(params)){
    firstActiveRow <- params$firstActiveRow
    freezePanes <- TRUE    
    if(length(firstActiveRow) != nSheets)
      firstActiveRow <- rep_len(firstActiveRow, length.out = nSheets)
  }
  
  firstActiveCol <- rep_len(1L, length.out = nSheets)
  if("firstActiveCol" %in% names(params)){
    firstActiveCol <- params$firstActiveCol
    freezePanes <- TRUE    
    if(length(firstActiveCol) != nSheets)
      firstActiveCol <- rep_len(firstActiveCol, length.out = nSheets)
  }
  
  firstRow <- rep_len(FALSE, length.out = nSheets)
  if("firstRow" %in% names(params)){
    firstRow <- params$firstRow
    freezePanes <- TRUE    
    if("list" %in% class(x) & length(firstRow) != nSheets)
      firstRow <- rep_len(firstRow, length.out = nSheets)
  }
  
  firstCol <- rep_len(FALSE, length.out = nSheets)
  if("firstCol" %in% names(params)){
    firstCol <- params$firstCol
    freezePanes <- TRUE    
    if("list" %in% class(x) & length(firstCol) != nSheets)
      firstCol <- rep_len(firstCol, length.out = nSheets)
  }
  
  if(freezePanes){
    for(i in 1:nSheets)
      freezePane(wb = wb,
                 sheet = i,
                 firstActiveRow = firstActiveRow[i],
                 firstActiveCol = firstActiveCol[i],
                 firstRow = firstRow[i],
                 firstCol = firstCol[i])
  }
  
  
  
  
  saveWorkbook(wb = wb, file = file, overwrite = overwrite)
  
  invisible(wb)
  
}

