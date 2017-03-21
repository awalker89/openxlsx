
#' @include class_definitions.R


Workbook$methods(setColWidths = function(sheet){
  
  sheet <- validateSheet(sheet)
  
  widths <- colWidths[[sheet]]
  hidden <- attr(colWidths[[sheet]], "hidden", exact = TRUE)
  if(length(hidden) != length(widths))
    hidden <- rep("0", length(widths))
  
  cols <- names(colWidths[[sheet]])
  
  autoColsInds <- widths %in% c("auto", "auto2")
  autoCols <- cols[autoColsInds]
  
  ## If any not auto
  if(any(!autoColsInds))
    widths[!autoColsInds] <- as.numeric(widths[!autoColsInds]) + 0.71
  
  ## If any auto
  if(length(autoCols) > 0){
    
    ## only run if data on worksheet
    if(worksheets[[sheet]]$sheet_data$n_elements == 0 ){
      
      missingAuto <- autoCols
      
    }else if(all(is.na(worksheets[[sheet]]$sheet_data$v))){
      
      missingAuto <- autoCols
      
    }else{
      
      ## First thing - get base font max character width
      baseFont <- getBaseFont()
      baseFontName <- unlist(baseFont$name, use.names = FALSE)
      if(is.null(baseFontName)){
        baseFontName <- "calibri"
      }else{
        baseFontName <- gsub(" ", ".", tolower(baseFontName), fixed = TRUE)
        if(!baseFontName %in% names(openxlsxFontSizeLookupTable)){
          baseFontName <- "calibri"
        }
      }
      
      baseFontSize <- unlist(baseFont$size, use.names = FALSE)
      if(is.null(baseFontSize)){
        baseFontSize <- 11
      }else{
        baseFontSize <- as.numeric(baseFontSize)
        baseFontSize <- ifelse(baseFontSize < 8, 8, ifelse(baseFontSize > 36, 36, baseFontSize))
      }
      
      baseFontCharWidth <- openxlsxFontSizeLookupTable[[baseFontName]][baseFontSize - 7]
      allCharWidths <- rep(baseFontCharWidth, worksheets[[sheet]]$sheet_data$n_elements)
      #########----------------------------------------------------------------
      
      ## get char widths for each style object
      if(length(styleObjects) > 0 & any(!is.na(worksheets[[sheet]]$sheet_data$style_id))){
        
        thisSheetName <- sheet_names[sheet]
        
        ## Calc font width for all styles on this worksheet
        styleIds <- worksheets[[sheet]]$sheet_data$style_id
        styObSubet <- styleObjects[sort(unique(styleIds))]
        stySubset <- lapply(styObSubet, "[[", "style") 
        
        ## loop through stlye objects assignin a charWidth else baseFontCharWidth
        styleCharWidths <- sapply(stySubset, get_style_max_char_width, USE.NAMES = FALSE)
        
        
        ## Now assign all cells a character width
        allCharWidths <- styleCharWidths[worksheets[[sheet]]$sheet_data$style_id]
        allCharWidths[is.na(allCharWidths)] <- baseFontCharWidth
        
      }
      
      ## Now check for columns that are auto2
      auto2Inds <- which(widths %in% "auto2")
      if(length(auto2Inds) > 0 & length(worksheets[[sheet]]$mergeCells) > 0){
        
        ## get cell merges
        merged_cells <- regmatches(worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))
        
        comps <- lapply(merged_cells, function(rectCoords) unlist(strsplit(rectCoords, split = ":")))  
        merge_cols <- lapply(comps, convertFromExcelRef)
        merge_cols <- lapply(merge_cols, function(x) x[x %in% cols[auto2Inds]]) ## subset to auto2Inds
        
        merge_rows <- lapply(comps, function(x) as.numeric(gsub("[A-Z]", "", x, perl = TRUE)))
        merge_rows <- merge_rows[sapply(merge_cols, length) > 0]
        merge_cols <- merge_cols[sapply(merge_cols, length) > 0]
        
        sd <- worksheets[[sheet]]$sheet_data
        
        if(length(merge_cols) > 0){
          
          all_merged_cells <- lapply(1:length(merge_cols), function(i) expand.grid("rows" =  min(merge_rows[[i]]):max(merge_rows[[i]]),
                                                                    "cols" =  min(merge_cols[[i]]):max(merge_cols[[i]]) ) )
          
          all_merged_cells <- do.call("rbind", all_merged_cells)
          
          ## only want the sheet data in here
          refs <- paste(all_merged_cells[[1]], all_merged_cells[[2]], sep = ",")
          existing_cells <- paste(worksheets[[sheet]]$sheet_data$rows, worksheets[[sheet]]$sheet_data$cols, sep = ",")
          keep <- which(!existing_cells %in% refs & !is.na(worksheets[[sheet]]$sheet_data$v))
          
          sd <- Sheet_Data$new()
          sd$cols <- worksheets[[sheet]]$sheet_data$cols[keep]
          sd$t <- worksheets[[sheet]]$sheet_data$t[keep]
          sd$v <- worksheets[[sheet]]$sheet_data$v[keep]
          sd$n_elements <- length(sd$cols)
          allCharWidths <- allCharWidths[keep]

        }else{
          sd <- worksheets[[sheet]]$sheet_data 
        }
        
        
      }else{
        sd <- worksheets[[sheet]]$sheet_data
      }
      
      ## Now that we have the max character width for the largest font on the page calculate the column widths
      calculatedWidths <- .Call("openxlsx_calc_column_widths", PACKAGE = "openxlsx",
                                sd,
                                unlist(sharedStrings, use.names = FALSE),
                                as.integer(autoCols),
                                allCharWidths,
                                baseFontCharWidth,
                                getOption("openxlsx.minWidth", 3),
                                getOption("openxlsx.maxWidth", 250))
      
      missingAuto <- autoCols[!autoCols %in% names(calculatedWidths)]
      widths[names(calculatedWidths)] <- calculatedWidths + 0.71
      
    }
    
    widths[missingAuto] <- 9.15
    
  }
  
  ## Calculate width of auto
  colNodes <- sprintf('<col min="%s" max="%s" width="%s" hidden = "%s" customWidth="1"/>', cols, cols, widths, hidden)
  
  ## Append new col widths XML to worksheets[[sheet]]$cols
  worksheets[[sheet]]$cols <<- append(worksheets[[sheet]]$cols, colNodes)
  
})




get_style_max_char_width <- function(thisStyle){
  
  fN <- unlist(thisStyle$fontName, use.names = FALSE)
  if(is.null(fN)){
    fN <- "calibri"
  }else{
    fN <- gsub(" ", ".", tolower(fN), fixed = TRUE)
    if(!fN %in% names(openxlsxFontSizeLookupTable)){
      fN <- "calibri"
    }
  }
  
  fS <- unlist(thisStyle$fontSize, use.names = FALSE)
  if(is.null(fS)){
    fS <- 11
  }else{
    fS <- as.numeric(fS)
    fS <- ifelse(fS < 8, 8, ifelse(fS > 36, 36, fS))
  }
  
  if("BOLD" %in% thisStyle$fontDecoration){
    styleMaxCharWidth <- openxlsxFontSizeLookupTableBold[[fN]][fS - 7]
  }else{
    styleMaxCharWidth <- openxlsxFontSizeLookupTable[[fN]][fS - 7]
  }
  
  return(styleMaxCharWidth)
  
}

