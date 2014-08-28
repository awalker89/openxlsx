


genBaseColStyle <- function(cc){
  
  colStyle <- createStyle()
  specialFormat <- TRUE
  
  if("date" %in% cc){
    colStyle <- createStyle(numFmt = "date")
    
  }else if(any(c("posixlt", "posixct", "posixt") %in% cc)){
    colStyle <- createStyle(numFmt = "longdate")
    
  }else if("currency" %in% cc){
    colStyle$numFmt <- list("numFmtId" = "164", "formatCode" = "&quot;$&quot;#,##0.00")
    
  }else if("accounting" %in% cc){
    colStyle$numFmt <- list("numFmtId" = "44")
    
  }else if("hyperlink" %in% cc){
    colStyle$fontDecoration <- "UNDERLINE"
    colStyle$fontColour <- list("rgb" = "FF0000FF")
    
  }else if("percentage" %in% cc){
    colStyle$numFmt <- list(numFmtId = "10")
    
  }else if("scientific" %in% cc){
    colStyle$numFmt <- list(numFmtId = "11")
    
  }else if("3" %in% cc | "comma" %in% cc){
    colStyle$numFmt <- list(numFmtId = "3")
    
  }else if("numeric" %in% cc & !grepl("[^0\\.,#]", getOption("openxlsx.numFmt", "GENERAL")) ){
    colStyle$numFmt <- list("numFmtId" = 9999, "formatCode" = getOption("openxlsx.numFmt"))
          
  }else{
    colStyle$numFmt <- list(numFmtId = "0")
    specialFormat <- FALSE
    
  }
  
  list("style" = colStyle,
       "specialFormat" = specialFormat)
  
  
}



Workbook$methods(surroundingBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType){
  
  sheet <- names(worksheets)[[sheet]]
  ## steps
  # get column class
  # get corresponding base style
  
  for(i in 1:nCol){
    
    tmp <- genBaseColStyle(colClasses[[i]])
    
    colStyle <- tmp$style
    specialFormat <- tmp$specialFormat
    
    ## create style objects
    sTop <- colStyle$copy()
    sMid <- colStyle$copy()
    sBot <- colStyle$copy()
    
    ## First column
    if(i == 1){
      
      
      if(nRow == 1 & nCol == 1){
        
        ## All
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour
        
        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour
        
        sTop$borderRight <- borderStyle
        sTop$borderRightColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol
               )))
        ))
        
      }else if(nCol == 1){
        
        ## Top
        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour
        
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        sTop$borderRight <- borderStyle
        sTop$borderRightColour <- borderColour
        
        ## Middle
        sMid$borderLeft <- borderStyle
        sMid$borderLeftColour <- borderColour
        
        sMid$borderRight <- borderStyle
        sMid$borderRightColour <- borderColour
        
        ## Bottom
        sBot$borderBottom <- borderStyle
        sBot$borderBottomColour <- borderColour
        
        sBot$borderLeft <- borderStyle
        sBot$borderLeftColour <- borderColour
        
        sBot$borderRight <- borderStyle
        sBot$borderRightColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol
               )))
        ))
        
        styleObjects <<- append(styleObjects, list(
          list(style = sMid,
               cells = list(list("sheet" = sheet,
                                 "rows" = (startRow + 1L):(startRow + nRow - 2L)   , #2nd -> 2nd to last
                                 "cols" = rep.int(startCol, nRow - 2L)
               )))
        ))
        
        styleObjects <<- append(styleObjects, list(
          list(style = sBot,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow + nRow - 1L,
                                 "cols" = startCol
               )))
        ))
        
        
      }else if(nRow == 1){
        
        ## All
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour
        
        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol
               )))
        ))
        
      }else{
        
        ## Top
        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour
        
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        ## Middle
        sMid$borderLeft <- borderStyle
        sMid$borderLeftColour <- borderColour
        
        ## Bottom
        sBot$borderLeft <- borderStyle
        sBot$borderLeftColour <- borderColour
        
        sBot$borderBottom <- borderStyle
        sBot$borderBottomColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol
               )))
        ))
        
        styleObjects <<- append(styleObjects, list(
          list(style = sMid,
               cells = list(list("sheet" = sheet,
                                 "rows" = (startRow + 1L):(startRow + nRow - 2L)   , #2nd -> 2nd to last
                                 "cols" = rep.int(startCol, nRow - 2L)
               )))
        ))
        
        styleObjects <<- append(styleObjects, list(
          list(style = sBot,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow + nRow - 1L,
                                 "cols" = startCol
               )))
        ))
        
        
      }
      
      
      
    }else if(i == nCol){
      
      if(nRow == 1){
        
        ## All
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour
        
        sTop$borderRight <- borderStyle
        sTop$borderRightColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol + nCol - 1L
               )))
        ))
        
        
      }else{
        
        ## Top
        sTop$borderRight <- borderStyle
        sTop$borderRightColour <- borderColour
        
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        ## Middle
        sMid$borderRight <- borderStyle
        sMid$borderRightColour <- borderColour
        
        ## Bottom
        sBot$borderRight <- borderStyle
        sBot$borderRightColour <- borderColour
        
        sBot$borderBottom <- borderStyle
        sBot$borderBottomColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol + nCol - 1L
               )))
        ))
        
        
        styleObjects <<- append(styleObjects, list(
          list(style = sMid,
               cells = list(list("sheet" = sheet,
                                 "rows" = (startRow + 1L):(startRow + nRow - 2L)   , #2nd -> 2nd to last
                                 "cols" = rep.int(startCol + nCol - 1L, nRow - 2L)
               )))
        ))
        
        
        
        styleObjects <<- append(styleObjects, list(
          list(style = sBot,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow + nRow - 1L,
                                 "cols" = startCol + nCol - 1L
               )))
        )) 
        
      }
      
    }else{  ## inside columns
      
      if(nRow == 1){
        
        ## Top      
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        ## Bottom      
        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol + i - 1L
               )))
        ))
        
      }else{
        
        ## Top      
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        ## Bottom      
        sBot$borderBottom <- borderStyle
        sBot$borderBottomColour <- borderColour
        
        styleObjects <<- append(styleObjects, list(
          list(style = sTop,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow,
                                 "cols" = startCol + i - 1L
               )))
        ))
        
        ## Middle
        if(specialFormat){
          styleObjects <<- append(styleObjects, list(
            list(style = sMid,
                 cells = list(list("sheet" = sheet,
                                   "rows" = (startRow + 1L):(startRow + nRow - 2L)   , #2nd -> 2nd to last
                                   "cols" = rep.int(startCol + i - 1L, nRow - 2L)
                 )))
          ))
        }
        
        styleObjects <<- append(styleObjects, list(
          list(style = sBot,
               cells = list(list("sheet" = sheet,
                                 "rows" = startRow + nRow - 1L,
                                 "cols" = startCol + i - 1L
               )))
        ))
      }
      
    } ## End of if(i == 1), i == NCol, else inside columns
  }## End of loop through columns
  
  
  invisible(0)
  
})









Workbook$methods(rowBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType){
  
  sheet <- names(worksheets)[[sheet]]
  ## steps
  # get column class
  # get corresponding base style
  
  for(i in 1:nCol){
    
    tmp <- genBaseColStyle(colClasses[[i]])
    sTop <- tmp$style
    
    ## First column
    if(i == 1){
      
      if (nCol == 1){
        
        ## All borders (rows and surrounding)
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour
        
        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour
        
        sTop$borderRight <- borderStyle
        sTop$borderRightColour <- borderColour
        
      }else{
        
        ## Top, Left, Bottom
        sTop$borderTop <- borderStyle
        sTop$borderTopColour <- borderColour
        
        sTop$borderBottom <- borderStyle
        sTop$borderBottomColour <- borderColour
        
        sTop$borderLeft <- borderStyle
        sTop$borderLeftColour <- borderColour
        
      }
      
      
    }else if(i == nCol){
      
      ## Top, Right, Bottom
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop$borderBottom <- borderStyle
      sTop$borderBottomColour <- borderColour
      
      sTop$borderRight <- borderStyle
      sTop$borderRightColour <- borderColour 
      
    }else{  ## inside columns
      
      ## Top, Middle, Bottom
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop$borderBottom <- borderStyle
      sTop$borderBottomColour <- borderColour
      
    } ## End of if(i == 1), i == NCol, else inside columns
    
    styleObjects <<- append(styleObjects, list(
      list(style = sTop,
           cells = list(list("sheet" = sheet,
                             "rows" = (startRow):(startRow + nRow - 1L),
                             "cols" = rep(startCol + i - 1L, nRow)
           )))
    ))
    
  }## End of loop through columns
  
  
  invisible(0)
  
})




Workbook$methods(columnBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType){
  
  sheet <- names(worksheets)[[sheet]]
  ## steps
  # get column class
  # get corresponding base style
  
  for(i in 1:nCol){
    
    tmp <- genBaseColStyle(colClasses[[i]])
    colStyle <- tmp$style
    specialFormat <- tmp$specialFormat
    
    ## create style objects
    sTop <- colStyle$copy()
    sMid <- colStyle$copy()
    sBot <- colStyle$copy()
    
    if(nRow == 1){
      
      ## Top
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop$borderBottom <- borderStyle
      sTop$borderBottomColour <- borderColour
      
      sTop$borderLeft <- borderStyle
      sTop$borderLeftColour <- borderColour
      
      sTop$borderRight <- borderStyle
      sTop$borderRightColour <- borderColour
      
      styleObjects <<- append(styleObjects, list(
        list(style = sTop,
             cells = list(list("sheet" = sheet,
                               "rows" = startRow,
                               "cols" = startCol + i - 1L
             )))
      ))
      
      
    }else{
      
      ## Top
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      sTop$borderLeft <- borderStyle
      sTop$borderLeftColour <- borderColour
      
      sTop$borderRight <- borderStyle
      sTop$borderRightColour <- borderColour
      
      ## Middle
      sMid$borderLeft <- borderStyle
      sMid$borderLeftColour <- borderColour
      
      sMid$borderRight <- borderStyle
      sMid$borderRightColour <- borderColour
      
      ## Bottom
      sBot$borderBottom <- borderStyle
      sBot$borderBottomColour <- borderColour
      
      sBot$borderLeft <- borderStyle
      sBot$borderLeftColour <- borderColour
      
      sBot$borderRight <- borderStyle
      sBot$borderRightColour <- borderColour
      
      colInd <- startCol + i - 1L
      
      styleObjects <<- append(styleObjects, list(
        list(style = sTop,
             cells = list(list("sheet" = sheet,
                               "rows" = startRow,
                               "cols" = colInd
             )))
      ))
      
      styleObjects <<- append(styleObjects, list(
        list(style = sMid,
             cells = list(list("sheet" = sheet,
                               "rows" = (startRow + 1L):(startRow + nRow - 2L),
                               "cols" = rep(colInd, nRow - 2L)
             )))
      ))
      
      
      styleObjects <<- append(styleObjects, list(
        list(style = sBot,
             cells = list(list("sheet" = sheet,
                               "rows" = startRow + nRow - 1L,
                               "cols" = colInd
             )))
      ))
      
    }
    
  }## End of loop through columns
  
  
  invisible(0)
  
})





Workbook$methods(allBorders = function(colClasses, sheet, startRow, startCol, nRow, nCol, borderColour, borderStyle, borderType){
  
  sheet <- names(worksheets)[[sheet]]
  ## steps
  # get column class
  # get corresponding base style
  
  for(i in 1:nCol){
    
    tmp <- genBaseColStyle(colClasses[[i]])
    sTop <- tmp$style
    
    ## All borders
    sTop$borderTop <- borderStyle
    sTop$borderTopColour <- borderColour
    
    sTop$borderBottom <- borderStyle
    sTop$borderBottomColour <- borderColour
    
    sTop$borderLeft <- borderStyle
    sTop$borderLeftColour <- borderColour
    
    sTop$borderRight <- borderStyle
    sTop$borderRightColour <- borderColour
    
    styleObjects <<- append(styleObjects, list(
      list(style = sTop,
           cells = list(list("sheet" = sheet,
                             "rows" = (startRow):(startRow + nRow - 1L),
                             "cols" = rep(startCol + i - 1L, nRow)
           )))
    ))
    
  }## End of loop through columns
  
  
  invisible(0)
  
})
