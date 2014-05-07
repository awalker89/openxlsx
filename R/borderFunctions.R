




surroundingBorders <- function(wb, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle){
  
  
  if(nRow == 1 & nCol == 1){
    
    ## single cell
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol, gridExpand=TRUE)
    
  }else if(nRow == 1){
    
    ## left
    addStyle(wb, sheet, createStyle(border="TopBottomLeft", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol, gridExpand=TRUE)
    
    ## right
    addStyle(wb, sheet, createStyle(border="TopBottomRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol + nCol - 1, gridExpand=TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="TopBottom", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols = (startCol+1):(startCol + nCol - 2), gridExpand=TRUE)   
    
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
    
    ## top left corner
    addStyle(wb, sheet, createStyle(border="TopLeft", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow, cols=startCol, gridExpand=TRUE)
    
    ## top right corner
    addStyle(wb, sheet, createStyle(border="TopRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow, cols=startCol + nCol - 1, gridExpand=TRUE)
    
    ## bottom left corner
    addStyle(wb, sheet, createStyle(border="BottomLeft", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow + nRow - 1, cols=startCol, gridExpand=TRUE)
    
    ## bottom right corner
    addStyle(wb, sheet, createStyle(border="BottomRight", borderColour=borderColour, borderStyle = borderStyle),
             rows=startRow + nRow - 1, cols=startCol + nCol - 1, gridExpand=TRUE)
    
    ## top
    addStyle(wb, sheet, createStyle(border="Top", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=(startCol+1):(startCol + nCol - 2), gridExpand=TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="Bottom", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow + nRow - 1, cols=(startCol+1):(startCol + nCol - 2), gridExpand=TRUE)
    
    ## left
    addStyle(wb, sheet, createStyle(border="Left", borderColour=borderColour, borderStyle = borderStyle),
             rows= (startRow + 1):(startRow + nRow - 2), cols=startCol, gridExpand=TRUE)
    
    ## right
    addStyle(wb, sheet, createStyle(border="Right", borderColour=borderColour, borderStyle = borderStyle),
             rows= (startRow + 1):(startRow + nRow - 2), cols=startCol + nCol - 1, gridExpand=TRUE)
    
  }
  
}


rowBorders <- function(wb, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle){
  
  if(nRow == 1 & nCol == 1){
    
    ## single cell
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol, gridExpand=TRUE)
    
  }else if(nRow == 1){
    
    ## left
    addStyle(wb, sheet, createStyle(border="TopBottomLeft", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol, gridExpand=TRUE)
    
    ## right
    addStyle(wb, sheet, createStyle(border="TopBottomRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols=startCol + nCol - 1, gridExpand=TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="TopBottom", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow, cols = (startCol+1):(startCol + nCol - 2), gridExpand=TRUE)   
    
  }else if(nCol == 1){
    
    ## single column
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow:(startRow + nRow -1), cols=startCol, gridExpand=TRUE)
    
  }else{
    
    ## left, leftTop, leftBottom
    addStyle(wb, sheet, createStyle(border="TopBottomLeft", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow:(startRow + nRow - 1), cols=startCol, gridExpand=TRUE)
    
    ## right, rightTop, rightBottom
    addStyle(wb, sheet, createStyle(border="TopBottomRight", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow:(startRow + nRow - 1), cols=startCol + nCol - 1, gridExpand=TRUE)
    
    ## all rows
    addStyle(wb, sheet, createStyle(border="TopBottom", borderColour=borderColour, borderStyle = borderStyle),
             rows= startRow:(startRow + nRow - 1), cols=(startCol+1):(startCol + nCol - 2), gridExpand = TRUE)
    
  }
}


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









