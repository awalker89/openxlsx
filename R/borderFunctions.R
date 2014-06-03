


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
    
    ## Column borders do not change depending on column
    ## Top
    if(borderType == "columns"){
      sTop$borderLeft <- borderStyle
      sTop$borderLeftColour <- borderColour
      
      sTop$borderRight <- borderStyle
      sTop$borderRightColour <- borderColour
      
      sTop$borderTop <- borderStyle
      sTop$borderTopColour <- borderColour
      
      ## Middle
      sMid$borderLeft <- borderStyle
      sMid$borderLeftColour <- borderColour
      
      sMid$borderRight <- borderStyle
      sMid$borderRightColour <- borderColour
      
      ## Bottom
      sBot$borderLeft <- borderStyle
      sBot$borderLeftColour <- borderColour
      
      sBot$borderRight <- borderStyle
      sBot$borderRightColour <- borderColour
      
      sBot$borderBottom <- borderStyle
      sBot$borderBottomColour <- borderColour
      
      sTop <- list(style = sTop,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow,
                                     "cols" = startCol + i - 1
                   )))
      
      sMid <- list(style = sMid,
                   cells = list(list("sheet" = sheet,
                                     "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                     "cols" = rep.int(startCol + i - 1, nRow - 2)
                   )))
      
      sBot <- list(style = sBot,
                   cells = list(list("sheet" = sheet,
                                     "rows" = startRow + nRow - 1,
                                     "cols" = startCol + i - 1
                   )))
      
      
    ## If borderType is Not "column"
    }else{
    
      
      ## First column
      if(i == 1){
        
        if(borderType == "surrounding"){
        
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
              
        
        }else{  ## borderType == "rows"
          
          ## Top, Left, Bottom
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour
          
          sTop$borderBottom <- borderStyle
          sTop$borderBottomColour <- borderColour
          
          sTop$borderLeft <- borderStyle
          sTop$borderLeftColour <- borderColour
          
          sMid <- NULL
          sBot <- NULL
          
        }
        
        if(!is.null(sTop))
          sTop <- list(style = sTop,
                       cells = list(list("sheet" = sheet,
                                         "rows" = startRow,
                                         "cols" = startCol
                       )))
        
        if(!is.null(sMid))
          sMid <- list(style = sMid,
                       cells = list(list("sheet" = sheet,
                                         "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                         "cols" = rep.int(startCol, nRow - 2)
                       )))
        
        if(!is.null(sBot))
          sBot <- list(style = sBot,
                       cells = list(list("sheet" = sheet,
                                         "rows" = startRow + nRow - 1,
                                         "cols" = startCol
                       )))
        
        
      }else if(i == nCol){
        
        if(borderType == "surrounding"){
          
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
          
        }else{  ## borderType == "rows"
          
          ## Top, Left, Bottom
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour
          
          sTop$borderBottom <- borderStyle
          sTop$borderBottomColour <- borderColour
          
          sTop$borderRight <- borderStyle
          sTop$borderRightColour <- borderColour
          
          sMid <- NULL
          sBot <- NULL
          
        }
        
        if(!is.null(sTop))
          sTop <- list(style = sTop,
                       cells = list(list("sheet" = sheet,
                                         "rows" = startRow,
                                         "cols" = startCol + nCol - 1
                       )))
        
        if(!is.null(sMid))
          sMid <- list(style = sMid,
                       cells = list(list("sheet" = sheet,
                                         "rows" = (startRow+1):(startRow + nRow -2)   , #2nd -> 2nd to last
                                         "cols" = rep.int(startCol + nCol - 1, nRow - 2)
                       )))
        
        if(!is.null(sBot))
          sBot <- list(style = sBot,
                       cells = list(list("sheet" = sheet,
                                         "rows" = startRow + nRow - 1,
                                         "cols" = startCol + nCol - 1
                       )))
        
  
      }else{  ## inside columns
        
        if(borderType == "surrounding"){
          
          ## Top      
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour
          
          ## Middle
          sMid <- NULL
          
          ## Bottom      
          sBot$borderBottom <- borderStyle
          sBot$borderBottomColour <- borderColour
          
          sTop <- list(style = sTop,
                       cells = list(list("sheet" = sheet,
                                         "rows" = startRow,
                                         "cols" = startCol + i - 1
                       )))
          
          sBot <- list(style = sBot,
                       cells = list(list("sheet" = sheet,
                                         "rows" = startRow + nRow - 1,
                                         "cols" = startCol + i - 1
                       )))
          
          
        }else{  ## borderType == "rows"
          
          ## Top, Middle, Bottom
          sTop$borderTop <- borderStyle
          sTop$borderTopColour <- borderColour
          
          sTop$borderBottom <- borderStyle
          sTop$borderBottomColour <- borderColour
          
          sMid <- NULL
          sBot <- NULL
  
          sTop <- list(style = sTop,
                       cells = list(list("sheet" = sheet,
                                         "rows" = (startRow):(startRow + nRow - 1),
                                         "cols" = startCol + i - 1
                       )))
          
        }
      
      } ## End of if(i == 1), i == NCol, else...
      
    } ## End of if(borderType == "columns") else...
      

    ## Append to styleObjects
    if(!is.null(sTop))
      styleObjects <<- append(styleObjects, list(sTop))
    
    if(!is.null(sMid))
      styleObjects <<- append(styleObjects, list(sMid))
    
    if(!is.null(sBot))
      styleObjects <<- append(styleObjects, list(sBot))
    
  }
  
  return(0)
  
})

