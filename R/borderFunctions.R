


Workbook$methods(makeBorderStyles = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType){

  sheet <- names(worksheets)[[sheet]]
  ## steps
  # get column class
  # get corresponding base style
  
  for(i in 1:nCol){
        
    cc <- colClasses[[i]]    
    specialFormat <- TRUE
    colStyle <- createStyle()
    
    if(any(c("date", "posixct", "posixt") %in% cc)){
      colStyle$numFmt <- list("numFmtId" = 14)
    }else if("currency" %in% cc){
      colStyle$numFmt <- list("numFmtId" = 164, "formatCode" = "&quot;$&quot;#,##0.00")
    }else if("accounting" %in% cc){
      colStyle$numFmt <- list("numFmtId" = 44)
    }else if("hyperlink" %in% cc){
      colStyle$fontDecoration <- "UNDERLINE"
      colStyle$fontColour <- list("rgb" = "FF0000FF")
    }else if("percentage" %in% cc){
      colStyle$numFmt <- list(numFmtId = 10)
    }else if("3" %in% cc){
      colStyle$numFmt <- list(numFmtId = 3)
    }else{
      colStyle$numFmt <- list(numFmtId = 0)
      specialFormat <- FALSE
    }
    
    
    ## create style objects
    sTop <- colStyle$copy()
    sMid <- colStyle$copy()
    sBot <- colStyle$copy()
    
    ## All columns require 3 styles unique styles
    ## First column
    if(i == 1){
      
      
      
      ## Top
      sTop$borderLeft <- borderStyle
      sTop$borderTop <- borderStyle
      sTop$borderLeftColour <- borderColour
      sTop$borderTopColour <- borderColour
      
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol
                   )))
      
      ## Middle
      sMid$borderLeft <- borderStyle
      sMid$borderLeftColour <- borderColour
      
      sMid <- list(style = sMid,
                   cells = list(list("sheet" = sheet,
                                     "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                     "cols" = rep.int(startCol, nRow - 2)
                   )))
      
      ## Bottom
      sBot$borderLeft <- borderStyle
      sBot$borderBottom <- borderStyle
      sBot$borderLeftColour <- borderColour
      sBot$borderBottomColour <- borderColour
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol
                   )))
      
      ## last column
    }else if(i == nCol){
      
      ## Top
      sTop$borderRight <- borderStyle
      sTop$borderTop <- borderStyle
      sTop$borderRightColour <- borderColour
      sTop$borderTopColour <- borderColour
      
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol + nCol - 1
                   )))
      
      ## Middle
      sMid$borderRight <- borderStyle
      sMid$borderRightColour <- borderColour
      
      sMid <- list(style = sMid,
                   cells = list(list("sheet" = sheet,
                                     "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                     "cols" = rep.int(startCol + nCol - 1, nRow - 2)
                   )))
      
      ## Bottom
      sBot$borderRight <- borderStyle
      sBot$borderBottom <- borderStyle
      sBot$borderRightColour <- borderColour
      sBot$borderBottomColour <- borderColour
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol + nCol - 1
                   )))
      
      
      
      
      
      
      ## center columns
    }else{
      
      ## startRow < i < nCol
      
      ## Top
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol + i - 1
                   )))
      
      ## Middle     
      if(specialFormat){
        sMid <- list(style = sMid,
                     cells = list(list("sheet" = sheet,
                                       "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                       "cols" = rep.int(startCol + i - 1, nRow - 2)
                     )))
      }else{
        sMid <- NULL
      }
      
      ## Bottom
      sBot$borderBottom <- borderStyle
      sBot$borderBottomColour <- borderColour
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol + i - 1
                   )))
      
      
    }
    
    ## Append to styleObjects
    styleObjects <<- append(styleObjects, list(sTop))
    styleObjects <<- append(styleObjects, list(sMid))
    styleObjects <<- append(styleObjects, list(sBot))
    
  }
  
  return(0)
  
})


Workbook$methods(rowBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle){
  
  sheet <- names(worksheets)[[sheet]]
  ## steps
  # get column class
  # get corresponding base style
  
  for(i in 1:nCol){
    
    cc <- colClasses[[i]]    
    specialFormat <- TRUE
    colStyle <- createStyle()
    
    if(any(c("date", "posixct", "posixt") %in% cc)){
      colStyle$numFmt <- list("numFmtId" = 14)
    }else if("currency" %in% cc){
      colStyle$numFmt <- list("numFmtId" = 164, "formatCode" = "&quot;$&quot;#,##0.00")
    }else if("accounting" %in% cc){
      colStyle$numFmt <- list("numFmtId" = 44)
    }else if("hyperlink" %in% cc){
      colStyle$fontDecoration <- "UNDERLINE"
      colStyle$fontColour <- list("rgb" = "FF0000FF")
    }else if("percentage" %in% cc){
      colStyle$numFmt <- list(numFmtId = 10)
    }else if("3" %in% cc){
      colStyle$numFmt <- list(numFmtId = 3)
    }else{
      colStyle$numFmt <- list(numFmtId = 0)
      specialFormat <- FALSE
    }
    
    ## create style objects
    sTop <- colStyle$copy()
    sMid <- colStyle$copy()
    sBot <- colStyle$copy()
    
    ## All columns require 3 styles unique styles
    ## First column
    if(i == 1){
      
      ## Top
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop$borderBottom <- borderStyle
      sTop$borderBottomColour <- borderColour
      
      sTop$borderLeft <- borderStyle
      sTop$borderLeftColour <- borderColour
        
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol
                   )))
      
      ## Middle
      sMid$borderLeft <- borderStyle
      sMid$borderLeftColour <- borderColour
      
      sMid <- list(style = sMid,
                   cells = list(list("sheet" = sheet,
                                     "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                     "cols" = rep.int(startCol, nRow - 2)
                   )))
      
      ## Bottom
      sBot$borderLeft <- borderStyle
      sBot$borderBottom <- borderStyle
      sBot$borderLeftColour <- borderColour
      sBot$borderBottomColour <- borderColour
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol
                   )))
      
      ## last column
    }else if(i == nCol){
      
      ## Top
      sTop$borderRight <- borderStyle
      sTop$borderTop <- borderStyle
      sTop$borderRightColour <- borderColour
      sTop$borderTopColour <- borderColour
      
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol + nCol - 1
                   )))
      
      ## Middle
      sMid$borderRight <- borderStyle
      sMid$borderRightColour <- borderColour
      
      sMid <- list(style = sMid,
                   cells = list(list("sheet" = sheet,
                                     "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                     "cols" = rep.int(startCol + nCol - 1, nRow - 2)
                   )))
      
      ## Bottom
      sBot$borderRight <- borderStyle
      sBot$borderBottom <- borderStyle
      sBot$borderRightColour <- borderColour
      sBot$borderBottomColour <- borderColour
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol + nCol - 1
                   )))
      
      
      
      
      
      
      ## center columns
    }else{
      
      ## startRow < i < nCol
      
      ## Top
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol + i - 1
                   )))
      
      ## Middle     
      if(specialFormat){
        sMid <- list(style = sMid,
                     cells = list(list("sheet" = sheet,
                                       "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                       "cols" = rep.int(startCol + i - 1, nRow - 2)
                     )))
      }else{
        sMid <- NULL
      }
      
      ## Bottom
      sBot$borderBottom <- borderStyle
      sBot$borderBottomColour <- borderColour
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol + i - 1
                   )))
      
      
    }
    
    ## Append to styleObjects
    styleObjects <<- append(styleObjects, list(sTop))
    styleObjects <<- append(styleObjects, list(sMid))
    styleObjects <<- append(styleObjects, list(sBot))
    
  }
  
  return(0)
  
})


colBorders <- function(wb, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle){
  
  if(nCol == 1 & nRow == 1){
    
    ## single cell
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow, cols=startCol, gridExpand=TRUE)
    
  }else if(nRow == 1){
    
    ## all
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow, cols=startCol:(startCol + nCol - 1), gridExpand=TRUE)
    
  }else if(nCol == 1){
    
    ## top
    addStyle(wb, sheet, createStyle(border="TopLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol, gridExpand=TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="BottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow+nRow-1, cols=startCol, gridExpand=TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="LeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= (startRow+1):(startRow+nRow-2), cols = startCol, gridExpand=TRUE)  
    
    
  }else{
    
    ## top
    addStyle(wb, sheet, createStyle(border="TopLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow, cols=startCol:(startCol + nCol - 1), gridExpand=TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="BottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow + nRow - 1, cols=startCol:(startCol + nCol - 1), gridExpand=TRUE)
    
    ## all other rows
    addStyle(wb, sheet, createStyle(border="LeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=(startRow+1):(startRow + nRow - 2), cols=startCol:(startCol + nCol - 1), gridExpand=TRUE)
    
  }
}









