


#' @name loadWorkbook 
#' @title Load an exisiting .xlsx file
#' @author Alexander Walker
#' @param xlsxFile A path to an existing .xlsx file
#' @description  loadWorkbook returns a workbook object conserving styles and 
#' formatting of the original .xlsx file. 
#' @return Workbook object. 
#' @export
#' @seealso \code{\link{removeWorksheet}}
#' @examples
#' ## load existing workbook from package folder
#' wb <- loadWorkbook(xlsxFile = system.file("loadExample.xlsx", package= "openxlsx"))
#' sheets(wb)  #list worksheets
#' wb ## view object
#' ## Add a worksheet
#' addWorksheet(wb, "A new worksheet")
#' 
#' ## update data in Sales worksheet
#' df <- data.frame("Orders" = round(runif(10)*200),
#'                  "TotalSales" = runif(10)*400)
#' writeData(wb, "Sales", df, xy=c("A", 2))
#' 
#' ## Save workbook
#' saveWorkbook(wb, "loadExample.xlsx", overwrite = TRUE)
loadWorkbook <- function(xlsxFile){

  
  if(!file.exists(xlsxFile))
    stop("File does not exist.")
  
  wb <- createWorkbook()
  
  
  ## create temp dir
  xmlDir <- paste0(tempdir(),  tempfile(tmpdir = ""), "_openxlsx_loadWorkbook")
  
  ## Unzip files to temp directory
  xmlFiles <- unzip(xlsxFile, exdir = xmlDir)
  
  .relsXML          <- xmlFiles[grepl("_rels/.rels$", xmlFiles, perl = TRUE)]
  drawingsXML       <- xmlFiles[grepl("drawing[0-9]+.xml$", xmlFiles, perl = TRUE)]
  worksheetsXML     <- xmlFiles[grepl("/worksheets/sheet[0-9]", xmlFiles, perl = TRUE)]
  appXML            <- xmlFiles[grepl("app.xml$", xmlFiles, perl = TRUE)]
  coreXML           <- xmlFiles[grepl("core.xml$", xmlFiles, perl = TRUE)]
  workbookXML       <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  stylesXML         <- xmlFiles[grepl("styles.xml$", xmlFiles, perl = TRUE)]
  sharedStringsXML  <- xmlFiles[grepl("sharedStrings.xml$", xmlFiles, perl = TRUE)]
  themeXML          <- xmlFiles[grepl("theme[0-9]+.xml$", xmlFiles, perl = TRUE)]
  drawingRelsXML    <- xmlFiles[grepl("drawing[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  sheetRelsXML      <- xmlFiles[grepl("sheet[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  media             <- xmlFiles[grepl("image[0-9]+.[a-z]+$", xmlFiles, perl = TRUE)]
  charts            <- xmlFiles[grepl("chart[0-9]+.[a-z]+$", xmlFiles, perl = TRUE)]
  tablesXML         <- xmlFiles[grepl("tables/table[0-9]+.xml$", xmlFiles, perl = TRUE)]
  tableRelsXML      <- xmlFiles[grepl("table[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  queryTablesXML    <- xmlFiles[grepl("queryTable[0-9]+.xml$", xmlFiles, perl = TRUE)]
  connectionsXML    <- xmlFiles[grepl("connections.xml$", xmlFiles, perl = TRUE)]
  extLinksXML       <- xmlFiles[grepl("externalLink[0-9]+.xml$", xmlFiles, perl = TRUE)]
  extLinksRelsXML   <- xmlFiles[grepl("externalLink[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  
  
  nSheets <- length(worksheetsXML)
  
  ## xl\
  ## xl\workbook
  if(length(workbookXML) > 0){
    workbook <- readLines(workbookXML, warn=FALSE, encoding="UTF-8")
    workbook <-  removeHeadTag(workbook)
    
    sheets <- unlist(regmatches(workbook, gregexpr("<sheet .*/sheets>", workbook, perl = TRUE)))
    sheetrId <- as.numeric(unlist(regmatches(sheets, gregexpr('(?<=r:id="rId)[0-9]+', sheets, perl = TRUE))))
    sheetId <- as.numeric(unlist(regmatches(sheets, gregexpr('(?<=sheetId=")[0-9]+', sheets, perl = TRUE))))
    sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
    
    calcPr <- .Call("openxlsx_getChildlessNode", workbook, "<calcPr ", PACKAGE = "openxlsx")
    if(length(calcPr) > 0)
      wb$workbook$calcPr <- calcPr
    
    ## Make sure sheets are in order
    sheetNames <- sheetNames[order(sheetrId)]
    sheetNames <- replaceXMLEntities(sheetNames)
    
    
    ## add worksheets to wb
    invisible(lapply(sheetNames, function(sheetName) wb$addWorksheet(sheetName)))
      
    ## defined Names
    dNames <- .Call("openxlsx_getNodes", workbook, "<definedNames>", PACKAGE = "openxlsx")
    if(length(dNames) > 0)
    wb$workbook$definedNames <- dNames
    
  }
  
  ## xl\sharedStrings
  if(length(sharedStringsXML) > 0){
    
    sharedStrings <- readLines(sharedStringsXML, warn = FALSE, encoding="UTF-8")
    sharedStrings <- removeHeadTag(sharedStrings)
      
    uniqueCount <- as.numeric(regmatches(sharedStrings, regexpr('(?<=uniqueCount=")[0-9]+', sharedStrings, perl = TRUE)))
    
    ## read in and get <si> nodes
    vals <- .Call("openxlsx_getNodes", sharedStrings, "<si>", PACKAGE = "openxlsx")
    attr(vals, "uniqueCount") <- uniqueCount
    
    wb$sharedStrings <- vals
    
  }
    
  ## xl\styles
  if(length(stylesXML) > 0){
    
    ## Build style objects from the styles XML
    styles <- readLines(stylesXML, warn = FALSE)
    styles <- removeHeadTag(styles)
    
    ## Indexed colours
    vals <- .Call("openxlsx_getNodes", styles, "<indexedColors>", PACKAGE = "openxlsx")
    if(length(vals) > 0)
      wb$styles$indexedColors <- paste0("<colors>", vals, "</colors>")
    
    ## dxf (don't need these, I don't think)
    dxf <- .Call("openxlsx_getNodes", styles, "<dxf>", PACKAGE = "openxlsx")
    if(length(dxf) > 0)
      wb$styles$dxfs <- paste(dxf, collapse = "")
    
    ## Number formats
    numFmts <- .Call("openxlsx_getChildlessNode", styles, "<numFmt ", PACKAGE = "openxlsx")
    numFmtFlag <- FALSE
    if(length(numFmts) > 0){
      
      numFmtsIds <- sapply(numFmts, function(x) .Call("openxlsx_getAttr", x, 'numFmtId="', PACKAGE = "openxlsx"), USE.NAMES = FALSE)
      formatCodes <- sapply(numFmts, function(x) .Call("openxlsx_getAttr", x, 'formatCode="', PACKAGE = "openxlsx"), USE.NAMES = FALSE)
      numFmts <-lapply(1:length(numFmts), function(i) list("numFmtId"= numFmtsIds[[i]], "formatCode"=formatCodes[[i]]))
      numFmtFlag <- TRUE
      
    }
    
    ## fonts will maintain, sz, color, name, family scheme
    fonts <- .Call("openxlsx_getNodes", styles, "<font>", PACKAGE = "openxlsx")
    wb$styles$fonts[[1]] <- fonts[[1]]
    fonts <- buildFontList(fonts)       
    
    fills <- .Call("openxlsx_getNodes", styles, "<fill>", PACKAGE = "openxlsx")
    fills <- lapply(fills, nodeAttributes)

    borders <- .Call("openxlsx_getNodes", styles, "<border>", PACKAGE = "openxlsx")     
    borders <- sapply(borders, buildBorder, USE.NAMES = FALSE)
    
    cellXfs <- .Call("openxlsx_getNodes", styles, "<cellXfs", PACKAGE = "openxlsx") 
    
    xf <- .Call("openxlsx_getChildlessNode", cellXfs, "<xf ", PACKAGE = "openxlsx")
    xfAttrs <- regmatches(xf, gregexpr('[a-zA-Z]+=".*?"', xf))
    xfNames <- lapply(xfAttrs, function(xfAttrs) regmatches(xfAttrs, regexpr('[a-zA-Z]+(?=\\=".*?")', xfAttrs, perl = TRUE)))
    xfVals <- lapply(xfAttrs, function(xfAttrs) regmatches(xfAttrs, regexpr('(?<=").*?(?=")', xfAttrs, perl = TRUE)))
    
    for(i in 1:length(xf))
      names(xfVals[[i]]) <- xfNames[[i]]
    
    styleObjects <- list()
    flag <- FALSE
    for(s in xfVals){

      style <- createStyle()
      if(any(s != "0")){
                
        if(s[["fontId"]] != "0"){
          thisFont <- fonts[[(as.numeric(s[["fontId"]])+1)]]

          if("sz" %in% names(thisFont))
            style$fontSize <- thisFont$sz
          
          if("name" %in% names(thisFont))
            style$fontName <- thisFont$name
          
          if("family" %in% names(thisFont))
            style$fontFamily <- thisFont$family
          
          if("color" %in% names(thisFont))
            style$fontColour <- thisFont$color
          
          if("scheme" %in% names(thisFont))
            style$fontScheme <- thisFont$scheme
          
          flags <- c("bold", "italic", "underline") %in% names(thisFont)
          if(any(flags)){
            style$fontDecoration <- NULL
            if(flags[[1]])
              style$fontDecoration <- append(style$fontDecoration, "BOLD")
            
            if(flags[[2]])
              style$fontDecoration <- append(style$fontDecoration, "ITALIC")
            
            if(flags[[3]])
              style$fontDecoration <- append(style$fontDecoration, "UNDERLINE")
          }
        }
        
        
        if(s[["numFmtId"]] != "0" & numFmtFlag){
          if(as.numeric(s[["numFmtId"]]) < 164){
            style$numFmt <- list(numFmtId = s[["numFmtId"]])
          }else{
            style$numFmt <- numFmts[[which(s[["numFmtId"]] == numFmtsIds)]]
          }
        }
        
        ## Border
        if(s[["borderId"]] != "0" & "applyBorder" %in% names(s)){
     
          thisBorder <- borders[[as.numeric(s[["borderId"]]) + 1]]
          
          if("borderLeft" %in% names(thisBorder)){
            style$borderLeft    <- thisBorder$borderLeft
            style$borderLeftColour <- thisBorder$borderLeftColour
          }
          
          if("borderRight" %in% names(thisBorder)){
            style$borderRight    <- thisBorder$borderRight
            style$borderRightColour <- thisBorder$borderRightColour
          }
          
          if("borderTop" %in% names(thisBorder)){
            style$borderTop    <- thisBorder$borderTop
            style$borderTopColour <- thisBorder$borderTopColour
          }
          
          if("borderBottom" %in% names(thisBorder)){
            style$borderBottom    <- thisBorder$borderBottom
            style$borderBottomColour <- thisBorder$borderBottomColour
          }
            
        }
        
        ## alignment
        if("applyAlignment" %in% names(s) & "horizontal" %in% names(s))
          style$halign <- s[["horizontal"]]
        
        if("applyAlignment" %in% names(s) & "vertical" %in% names(s))
          style$valign <- s[["vertical"]]  
        
        if("applyAlignment" %in% names(s) & "textRotation" %in% names(s))
          style$textRotation <- s[["textRotation"]]
          
        ## wrap text
        if("wrapText" %in% names(s)){
          if(s[["wrapText"]] %in% c("1", "true"))
            style$wrapText <- TRUE
        }
        
        if(s[["fillId"]] != "0" && "applyFill" %in% names(s)){
          
          fillId <- as.numeric(s[["fillId"]]) + 1
          
          tmpFg <- fills[[fillId]]$fgColor
          tmpBg <- fills[[fillId]]$bgColor
          
          if(!is.null(tmpFg))
            style$fill$fillFg <- tmpFg
          
          if(!is.null(tmpFg))
            style$fill$fillBg <- tmpBg
          
        }
        
        
      } ## end if !all(s == "0)
      
      if(flag)
        styleObjects <- append(styleObjects , list(list(style = style, cells = list())))
      
      flag <- TRUE
      
    }  ## end of for loop through styles s in ...
    
    if(length(styleObjects) > 0)
      wb$styleObjects <- styleObjects
    
  }  ## end of if length(styleXML) > 0
  
  
  
  ## xl\media
  if(length(media) > 0){
    mediaNames <- regmatches(media, regexpr("image[0-9]\\.[a-z]+$", media))
    fileTypes <- unique(gsub("image[0-9]\\.", "", mediaNames))
    wb$Content_Types <- c(sprintf('<Default Extension="%s" ContentType="image/%s"/>', fileTypes, fileTypes), wb$Content_Types) 
    names(media) <- mediaNames
    wb$media <- media
  }
  
  ## xl\chart
  if(length(charts) > 0){
    chartNames <- regmatches(charts, regexpr("chart[0-9]\\.[a-z]+$", charts))
    names(charts) <- chartNames
    wb$charts <- charts
    wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/chart%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', 1:length(charts)))
  }
  
  ## xl\theme
  if(length(themeXML) > 0)
    wb$theme <- removeHeadTag(paste(unlist(lapply(sort(themeXML)[[1]], function(x) readLines(x, warn = FALSE, encoding = "UTF-8"))), collapse = ""))
  
  
  
  
  ## tables
  if(length(tablesXML) > 0){
    wb$tables <- lapply(sort(tablesXML), function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
    wb$Content_Types <- c(wb$Content_Types, 
                          sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', 1:length(wb$tables)+2))   
  }
  
  ## queryTables
  if(length(queryTablesXML) > 0){
    
    wb$queryTables <- lapply(sort(queryTablesXML), function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
    wb$Content_Types <- c(wb$Content_Types, 
                          sprintf('<Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>', 1:length(queryTablesXML)))   
  }
  
  
  ## connections
  if(length(connectionsXML) > 0){
    wb$connections <- removeHeadTag(.Call("openxlsx_cppReadFile", connectionsXML, PACKAGE = "openxlsx"))
    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>')
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>')
  }

  ## externalLinks
  if(length(extLinksXML)){
    wb$externalLinks <- lapply(sort(extLinksXML), function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
    
    wb$Content_Types <-c(wb$Content_Types, 
        sprintf('<Override PartName="/xl/externalLinks/externalLink%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>', 1:length(extLinksXML)))

    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf('<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>',
                                                            1:length(extLinksXML)))
  }
    
    
  ## externalLinksRels
  if(length(extLinksRelsXML))
    wb$externalLinksRels <- lapply(sort(extLinksRelsXML), function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
    
  ## xl\worksheets
  ws <- lapply(sort(worksheetsXML), function(x) readLines(x, warn = FALSE, encoding = "UTF-8"))
  ws <- lapply(ws, removeHeadTag)
  wsTemp <- lapply(ws, function(x) strsplit(x, split = "<sheetData>")[[1]])
    
  noData <- grepl("<sheetData/>", ws)
  if(any(noData))
    wsTemp[noData] <- lapply(ws[noData], function(x) strsplit(x, split = "<sheetData/>")[[1]])                            
  
  notSplit <- sapply(wsTemp, length) == 1
  if(any(notSplit)){
    wsTemp[notSplit] <- list(wsTemp[notSplit], wsTemp[notSplit])
  }
  
  wsData <- lapply(wsTemp, "[[", 1)
  sheetData <- lapply(wsTemp, "[[", 2)
  
  rows <- lapply(sheetData, function(x) strsplit(x, split = "<row ")[[1]])
  rows <- lapply(rows, function(x) x[grepl('r="', x)])
  
  cells <- lapply(sheetData, function(x) unlist(strsplit(x, split = "<c ")))
  cells <- lapply(cells, function(x) paste0("<c ", x[2:length(x)]))
  cells[noData] <- "No Data"
  
  rowNumbers <- styleRefs <- s <- replicate(NULL, n = nSheets)

  for(i in 1:length(cells)){
    
    if(!noData[[i]]){
    
      rowNumbers[[i]] <- .Call("openxlsx_getRefs", rows[[i]], 1, PACKAGE = "openxlsx")
      
      r <- .Call("openxlsx_getRefs", cells[[i]], 1, PACKAGE = "openxlsx")
      cells[[i]] <- .Call("openxlsx_getCells", cells[[i]], PACKAGE = "openxlsx")
      
      v <- .Call("openxlsx_getVals", cells[[i]], PACKAGE = "openxlsx")
      t <- .Call("openxlsx_getCellTypes", cells[[i]], PACKAGE = "openxlsx")
      f <- .Call("openxlsx_getFunction", cells[[i]], PACKAGE = "openxlsx")
      
      ## XML replacements
      v <- gsub('"', "&quot;", v)
      v <- gsub('&', "&amp;", v)
      v <- gsub("'", "&apos;", v)
      v <- gsub('<', "&lt;", v)
      v <- gsub('>', "&gt;", v)
      
      if(length(v) > 0)
        t[is.na(v)] <- as.character(NA)
      
      ## sheetData
      tmp <- .Call("openxlsx_buildLoadCellList", r , t , v, f, PACKAGE="openxlsx")
      names(tmp) <- gsub("[A-Z]", "", r) 
      wb$sheetData[[i]] <- tmp
      wb$dataCount[[i]] <- 1
      
      cells[[i]] <- cells[[i]][which(grepl(' s="', cells[[i]]))]
      s[[i]] <- .Call("openxlsx_getCellStyles", cells[[i]], PACKAGE = "openxlsx")
      styleRefs[[i]] <- .Call("openxlsx_getRefs", cells[[i]], 1, PACKAGE = "openxlsx")
    
    }
    ## row heights
    if(length(rows[[i]]) > 0){
      customRowHeights <- .Call("openxlsx_getAttr", rows[[i]], ' ht="', PACKAGE = "openxlsx")
      if(any(!is.na(customRowHeights)))
        setRowHeights(wb, i, rows = rowNumbers[[i]][!is.na(customRowHeights)], heights = as.numeric(customRowHeights[!is.na(customRowHeights)]))
    }
    
    ## Custom col widths
    cols <- .Call("openxlsx_getChildlessNode", wsData[[i]], "<col ")
    cols <- cols[grepl("customWidth", cols)]
    if(length(cols) > 0){
      
      mins <- as.numeric(.Call("openxlsx_getAttr", cols, 'min="', PACKAGE = "openxlsx"))
      max <- as.numeric(.Call("openxlsx_getAttr", cols, 'max="', PACKAGE = "openxlsx"))
      widths <-  .Call("openxlsx_getAttr", cols, 'width="', PACKAGE = "openxlsx") 
      
      if(any(mins != max)){
        mins <- lapply(1:length(mins), function(i) seq(mins[[i]], max[[i]]))
        widths <- unlist(lapply(1:length(widths), function(i) rep(widths[[i]], length(mins[[i]]))))  
      }  
      setColWidths(wb, i, cols = unlist(mins), widths = as.numeric(widths)-0.71)
    }
    
    ## auto filters
    autoFilter <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<autoFilter ")
    if(length(autoFilter) > 0){
      autoFilter <- paste0("<", unlist(strsplit(autoFilter, split = "<"))[[2]])
      if(!grepl("/>$", autoFilter))
        autoFilter <- gsub(">$", "/>", autoFilter)
      
      wb$worksheets[[i]]$autoFilter <- autoFilter
    }
    
    ## hyperlinks
    hyperlinks <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<hyperlink ")
    if(length(hyperlinks) > 0)
      wb$hyperlinks[[i]] <- .Call("openxlsx_getHyperlinkRefs", hyperlinks, 1, PACKAGE = "openxlsx")
    
    ## pageMargins
    pageMargins <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<pageMargins ")
    if(length(pageMargins) > 0)
      wb$worksheets[[i]]$pageMargins <- pageMargins
    
    ## pageSetup
    pageSetup <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<pageSetup ")
    if(length(pageSetup) > 0)
      wb$worksheets[[i]]$pageSetup <- gsub('r:id="rId[0-9]+"', 'r:id="rId2"', pageSetup)
    
    ## tableParts
    tableParts <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<tablePart ")
    if(length(tableParts) > 0)
      wb$worksheets[[i]]$tableParts <- sprintf('<tablePart r:id="rId%s"/>', 1:length(tableParts)+1)
    
    ## Merge Cells
    merges <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<mergeCell ")
    if(length(merges) > 0)
      wb$worksheets[[i]]$mergeCells <- merges
    
    ## freeze pane
    pane <- .Call("openxlsx_getChildlessNode", wsData[[i]], "<pane ")
    if(length(pane) > 0)
      wb$freezePane[[i]] <- pane
    
    ## Drawing
    drawingId <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<drawing ")
    if(length(drawingId) == 0)
      drawingId <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<legacyDrawing ")
    
    if(length(drawingId) > 0)
      wb$worksheets[[i]]$drawing <- gsub('rId.+"', 'rId1"', drawingId)

  }

  
  ## styles
  for(i in 1:nSheets){
  
    ## Sheet Data
    if(length(s[[i]]) > 0){
    
      ## write style coords to styleObjects (will be in order)
      sx <- s[[i]]
      
      if(any(!is.na(sx))){
        sRef <- lapply(as.numeric(unique(sx[!is.na(sx)])), function(styleInd){ 
          
          if(styleInd > 0){
            sRef <- styleRefs[[i]][sx == styleInd & !is.na(sx)]        
            rows <- as.numeric(gsub("[^0-9]", "", sRef))
            cols <- .Call("openxlsx_RcppConvertFromExcelRef", sRef)
                                                
            wb$styleObjects[[styleInd]]$cells <- append(wb$styleObjects[[styleInd]]$cells, 
                                                        list(list(sheet = wb$getSheetName(1:nSheets)[[i]],
                                                                  rows = rows,
                                                                  cols = cols)))
          }
        })
      }
    }
  }
  
  ############## worksheets complete
  
  ## Next sheetRels to see which drawings_rels belongs to which sheet
  
  if(length(sheetRelsXML) > 0){
    sheetRelsXML <- sort(sheetRelsXML)
    sheetNumber <- as.numeric(regmatches(sheetRelsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", sheetRelsXML, perl = TRUE)))
    

    xml <- lapply(sheetRelsXML, readLines, warn = FALSE)
    xml <- unlist(lapply(xml, removeHeadTag))
    xml <- gsub("<Relationships .*?>", "", xml)
    xml <- gsub("</Relationships>", "", xml)
    xml <- lapply(xml, function(x) .Call("openxlsx_getChildlessNode", x, "<Relationship ", PACKAGE="openxlsx"))
    
    
    if(length(tablesXML) > 0){
  
      tables <- lapply(xml, function(x) as.numeric(regmatches(x, regexpr("(?<=table)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      if(length(unlist(tables)) > 0){    
        ## get the tables that belong to each worksheet and create a worksheets_rels for each
        tCount <- 2 ## table r:Ids start at 2
        for(i in 1:length(tables)){
          k <- 1:length(tables[[i]]) + tCount
          wb$worksheets_rels[[sheetNumber[i]]] <- unlist(c(wb$worksheets_rels[[sheetNumber[i]]],
                                                           sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',  k,  k)))
          
          
          wb$worksheets[[sheetNumber[i]]]$tableParts <- sprintf("<tablePart r:id=\"rId%s\"/>", k)
          tCount <- tCount + length(k)
        }
        
        wb$tables <- wb$tables[order(unlist(tables))]
        
        ## relabel ids
        for(i in 1:length(wb$tables)){
          newId <- sprintf(' id="%s" ', i+2)
          wb$tables[[i]] <- sub(' id="[0-9]+" ' ,newId, wb$tables[[i]])
        }
      }
    } ## if(length(tablesXML) > 0)
    
    
    hlinks <- lapply(xml, function(x) x[grepl("hyperlink", x) & grepl("External", x)])
    hlinksInds <- which(sapply(hlinks, length) > 0)
    
    if(length(hlinksInds) > 0){
    
      hlinks <- hlinks[hlinksInds]
      for(i in hlinksInds){
        targets <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
        targets <- replaceXMLEntities(gsub('"$', "", targets))
        
        names(wb$hyperlinks[[sheetNumber[[i]]]]) <- targets  
      }
      
    }
    
  } 

  
  ## sheet i has sheet_rels i
  ## sheet rels i will link to drawing.xml j via sheetRels Target
  ## drawing.xml j has a drawing.rels j
  ## drwaings.rels j will link to a image in media
  
  ## find which sheet links to which drawing.xml
  ## look in corresping drawing rels for the image
  
  if(length(drawingRelsXML) > 0){
    
    xml <- lapply(sort(drawingRelsXML), readLines, warn = FALSE)
    xml <- unlist(lapply(xml, removeHeadTag))
    xml <- gsub("<Relationships .*?>", "", xml)
    xml <- gsub("</Relationships>", "", xml)
    
    wb$drawings_rels[sheetNumber] <- xml

  }

  ## table rels
  if(length(tableRelsXML) > 0){
    
    tableRelsXML <- sort(tableRelsXML)
    wb$tables.xml.rels <- character(length=length(tablesXML))

    ## which sheet does it belong to
    inds <- as.numeric(regmatches(tableRelsXML, regexpr("(?<=table)[0-9]+(?=\\.xml.rels$)", tableRelsXML, perl = TRUE)))
    xml <- sapply(tableRelsXML, function(x) .Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"), USE.NAMES = FALSE)
    xml <- sapply(xml, removeHeadTag, USE.NAMES = FALSE)
    wb$tables.xml.rels[inds] <- xml
    
  }else if(length(tablesXML) > 0){
    wb$tables.xml.rels <- rep("", length(tablesXML))
  }
  
  ## worksheet i has drawing i via rId:1
  
  
  ## Drawings
  ## remove head opend/close tags
  ## split each drawing
  
  if(length(drawingsXML > 0)){

    xml <- lapply(sort(drawingsXML), readLines, warn = FALSE)  
    xml <- unlist(lapply(xml, removeHeadTag))
    xml <- gsub("<xdr:wsDr .*?>", "", xml)
    xml <- gsub("</xdr:wsDr>", "", xml)
    
    ## split at one/two cell Anchor
    xml <- regmatches(xml, gregexpr("<xdr:...CellAnchor.*?</xdr:...CellAnchor>", xml))
    wb$drawings[sheetNumber] <- xml
  } 
  
  ## calc chain
#   if(length(calcChainXML) > 0){
#     
#     xml <- readLines(calcChainXML, warn = FALSE)
#     xml <- removeHeadTag(xml)
#     xml <- gsub("<calcChain .*?>", "", xml)
#     xml <- gsub("</calcChain>", "", xml)  
#     xml <- .Call("openxlsx_getChildlessNode", xml, "<c ")
#     
#     ## need to reset all i attributes to 1:nSheets
#     replacements <- paste0('i="', 1:nSheets, '"')
#     original <- paste0('i="', sheetId, '"')
#     
#     for(i in 1:length(replacements))
#       xml <- gsub(original[[i]], replacements[[i]], xml)
#       
#     wb$calcChain <- xml
#     wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>')
#     wb$workbook.xml.rels <- c(wb$workbook.xml.rels, '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>')
#   }
#   
  
  return(wb)
  
}


getAttrs <- function(xml, tag){
  
  x <- lapply(xml, function(x) .Call("openxlsx_getChildlessNode", x, tag, PACKAGE = "openxlsx"))
  x[sapply(x, length) == 0] <- ""
  a <- lapply(x, function(x) regmatches(x, regexpr('[a-zA-Z]+=".*?"', x)))
  
  names = lapply(a, function(xml) regmatches(xml, regexpr('[a-zA-Z]+(?=\\=".*?")', xml, perl = TRUE)))
  vals =  lapply(a, function(xml) regmatches(xml, regexpr('(?<=").*?(?=")', xml, perl = TRUE)))
  names(vals) <- names
  return(vals)
  
}


buildFontList <- function(fonts){
  
  sz <- getAttrs(fonts, "<sz ")
  colour <- getAttrs(fonts, "<color ")
  name <- getAttrs(fonts, "<name ")
  family <- getAttrs(fonts, "<family ")
  scheme <- getAttrs(fonts, "<scheme ")
  
  italic <- lapply(fonts, function(x) .Call("openxlsx_getChildlessNode", x, "<i", PACKAGE = "openxlsx"))
  bold <- lapply(fonts, function(x) .Call("openxlsx_getChildlessNode", x, "<b", PACKAGE = "openxlsx"))
  underline <- lapply(fonts, function(x) .Call("openxlsx_getChildlessNode", x, "<u", PACKAGE = "openxlsx"))
  
  ## Build font objects
  ft <- replicate(list(), n=length(fonts))
  for(i in 1:length(fonts)){
    
    f <- NULL
    nms <- NULL
    if(length(unlist(sz[i])) > 0){
      f <- c(f, sz[i])
      nms <- c(nms, "sz")     
    }
    
    if(length(unlist(colour[i])) > 0){
      f <- c(f, colour[i])
      nms <- c(nms, "color")  
    }
    
    if(length(unlist(name[i])) > 0){
      f <- c(f, name[i])
      nms <- c(nms, "name") 
    }
    
    if(length(unlist(family[i])) > 0){
      f <- c(f, family[i])
      nms <- c(nms, "family") 
    }
    
    if(length(unlist(scheme[i])) > 0){
      f <- c(f, scheme[i])
      nms <- c(nms, "scheme") 
    }
    
    if(length(italic[[i]]) > 0){
      f <- c(f, "italic")
      nms <- c(nms, "italic") 
    }
    
    if(length(bold[[i]]) > 0){
      f <- c(f, "bold")
      nms <- c(nms, "bold") 
    }
    
    if(length(underline[[i]]) > 0){
      f <- c(f, "underline")
      nms <- c(nms, "underline") 
    }
    
    f <- lapply(1:length(f), function(i) unlist(f[i]))
    names(f) <- nms
    
    ft[[i]] <- f
    
  }
  
  ft
  
}


nodeAttributes <- function(x){
  
  
  x <- paste0("<", unlist(strsplit(x, split = "<")))
  x <- x[grepl("<bgColor|<fgColor", x)]
  
  if(length(x) == 0)
    return("")
  
  attrs <- regmatches(x, gregexpr(' [a-zA-Z]+="[^"]*', x, perl =  TRUE))
  tags <- regmatches(x, gregexpr('<[a-zA-Z]+ ', x, perl =  TRUE))
  tags <- lapply(tags, gsub, pattern = "<| ", replacement = "")  
  attrs <- lapply(attrs, gsub, pattern = '"', replacement = "")      
  
  attrs <- lapply(attrs, strsplit, split = "=")
  for(i in 1:length(attrs)){
    nms <- lapply(attrs[[i]], "[[", 1)
    vals <- lapply(attrs[[i]], "[[", 2)
    a <- unlist(vals)
    names(a) <- unlist(nms)
    attrs[[i]] <- a
  }
  
  names(attrs) <- unlist(tags)
  
  attrs
}


buildBorder <- function(x){
  
  ## gets all borders that have children
  x <- unlist(lapply(c("<left", "<right", "<top", "<bottom"), function(tag) .Call("openxlsx_getNodes", x, tag, PACKAGE = "openxlsx")))
  if(length(x) == 0)
    return(NULL)
    
  sides <- c("TOP", "BOTTOM", "LEFT", "RIGHT")
  sideBorder <- character(length=length(x))
  for(i in 1:length(x)){
    tmp <- sides[sapply(sides, function(s) grepl(s, x[[i]], ignore.case = TRUE))]
    if(length(tmp) > 1) tmp <- tmp[[1]]
    if(length(tmp) == 1)
      sideBorder[[i]] <- tmp
  }
  
  sideBorder <- sideBorder[sideBorder != ""]
  x <- x[sideBorder != ""]
  if(length(sideBorder) == 0)
    return(NULL)
  
  
  ## style
  weight <- gsub('style=|"', "", regmatches(x, regexpr('style="[a-z]+"', x, perl = TRUE)))
    
  ## Colours
  cols <- replicate(n = length(sideBorder), list(rgb = "FF000000"))
  colNodes <- unlist(sapply(x, function(xml) .Call("openxlsx_getChildlessNode", xml, "<color", PACKAGE = "openxlsx"), USE.NAMES = FALSE))

  if(length(colNodes) > 0){
    attrs <- regmatches(colNodes, regexpr('(theme|indexed|rgb)=".+"', colNodes))
  }else{
    attrs <- NULL
  }
    
  if(length(attrs) != length(x)){
   return(
     list("borders" = paste(sideBorder, collapse = ""),
          "colour" = cols)
     ) 
  }
  
  attrs <- strsplit(attrs, split = "=")
  cols <- sapply(attrs, function(attr){
    y <- list(gsub('"', "", attr[[2]]))
    names(y) <- gsub(" ", "", attr[[1]])
    y
  })
  
  ## sideBorder & cols
  style <- list()
  
  if("LEFT" %in% sideBorder){
    style$borderLeft <- weight[which(sideBorder == "LEFT")]
    style$borderLeftColour <- cols[which(sideBorder == "LEFT")]
  }
  
  if("RIGHT" %in% sideBorder){
    style$borderRight <- weight[which(sideBorder == "RIGHT")]
    style$borderRightColour <- cols[which(sideBorder == "RIGHT")]
  }
  
  if("TOP" %in% sideBorder){
    style$borderTop <- weight[which(sideBorder == "TOP")]
    style$borderTopColour <- cols[which(sideBorder == "TOP")]
  }
  
  if("BOTTOM" %in% sideBorder){
    style$borderBottom <- weight[which(sideBorder == "BOTTOM")]
    style$borderBottomColour <- cols[which(sideBorder == "BOTTOM")]
  }
  
  return(style)
}




