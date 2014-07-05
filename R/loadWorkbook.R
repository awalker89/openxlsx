


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
    
    contentNodes <- sprintf('<Default Extension="%s" ContentType="image/%s"/>', fileTypes, fileTypes)
    contentNodes[fileTypes == "emf"] <- '<Default Extension="emf" ContentType="image/x-emf"/>'
    
    wb$Content_Types <- c(contentNodes, wb$Content_Types) 
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
    tablesXML <- tablesXML[order(nchar(tablesXML), tablesXML)]
    wb$tables <- lapply(tablesXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
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
    
  
  ##*----------------------------------------------------------------------------------------------*##
  ### BEGIN READING IN WORKSHEET DATA
  ##*----------------------------------------------------------------------------------------------*##
  
  ## xl\worksheets
  wsNumber <- as.numeric(regmatches(worksheetsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", worksheetsXML, perl = TRUE)))
  worksheetsXML <- worksheetsXML[order(wsNumber)]
  
  ws <- lapply(worksheetsXML, function(x) readLines(x, warn = FALSE, encoding = "UTF-8"))
  ws <- lapply(ws, removeHeadTag)
  wsTemp <- lapply(ws, function(x) strsplit(x, split = "<sheetData>")[[1]])
    
  ## If there need to split at sheetData tag
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
      v <- replaceIllegalCharacters(v)
      
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
    autoFilter <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<autoFilter ", PACKAGE = "openxlsx")
    if(length(autoFilter) > 0){
      autoFilter <- paste0("<", unlist(strsplit(autoFilter, split = "<"))[[2]])
      if(!grepl("/>$", autoFilter))
        autoFilter <- gsub(">$", "/>", autoFilter)
      
      wb$worksheets[[i]]$autoFilter <- autoFilter
    }
    
    ## hyperlinks
    hyperlinks <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<hyperlink ", PACKAGE = "openxlsx")
    if(length(hyperlinks) > 0)
      wb$hyperlinks[[i]] <- .Call("openxlsx_getHyperlinkRefs", hyperlinks, 1, PACKAGE = "openxlsx")
    
    ## pageMargins
    pageMargins <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<pageMargins ", PACKAGE = "openxlsx")
    if(length(pageMargins) > 0)
      wb$worksheets[[i]]$pageMargins <- pageMargins
    
    ## pageSetup
    pageSetup <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<pageSetup ", PACKAGE = "openxlsx")
    if(length(pageSetup) > 0)
      wb$worksheets[[i]]$pageSetup <- gsub('r:id="rId[0-9]+"', 'r:id="rId2"', pageSetup)
    
    ## tableParts
    tableParts <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<tablePart ", PACKAGE = "openxlsx")
    if(length(tableParts) > 0)
      wb$worksheets[[i]]$tableParts <- sprintf('<tablePart r:id="rId%s"/>', 1:length(tableParts)+1)
    
    ## Merge Cells
    merges <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<mergeCell ", PACKAGE = "openxlsx")
    if(length(merges) > 0)
      wb$worksheets[[i]]$mergeCells <- merges
    
    ## freeze pane
    pane <- .Call("openxlsx_getChildlessNode", wsData[[i]], "<pane ", PACKAGE = "openxlsx")
    if(length(pane) > 0)
      wb$freezePane[[i]] <- pane
    
    ## sheetView
    sheetViews <- .Call("openxlsx_getNodes", wsData[[i]], "<sheetViews>", PACKAGE = "openxlsx")
    if(length(sheetViews) > 0)
      wb$worksheets[[i]]$sheetViews <- sheetViews
    
    ## Drawing
    drawingId <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<drawing ", PACKAGE = "openxlsx")
    if(length(drawingId) == 0)
      drawingId <- .Call("openxlsx_getChildlessNode", sheetData[[i]], "<legacyDrawing ", PACKAGE = "openxlsx")
    
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
  
  ##*----------------------------------------------------------------------------------------------*##
  ### READING IN WORKSHEET DATA COMPLETE
  ##*----------------------------------------------------------------------------------------------*##
  
  ## Next sheetRels to see which drawings_rels belongs to which sheet
  
  if(length(sheetRelsXML) > 0){
    
    sheetRelsXML <- sheetRelsXML[order(nchar(sheetRelsXML), sheetRelsXML)]
    sheetNumber <- as.numeric(regmatches(sheetRelsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", sheetRelsXML, perl = TRUE)))

    xml <- lapply(sheetRelsXML, readLines, warn = FALSE)
    xml <- unlist(lapply(xml, removeHeadTag))
    xml <- gsub("<Relationships .*?>", "", xml)
    xml <- gsub("</Relationships>", "", xml)
    xml <- lapply(xml, function(x) .Call("openxlsx_getChildlessNode", x, "<Relationship ", PACKAGE="openxlsx"))
    
    ## tables
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
    
    ## hyperlinks
    hlinks <- lapply(xml, function(x) x[grepl("hyperlink", x) & grepl("External", x)])
    hlinksInds <- which(sapply(hlinks, length) > 0)
    
    if(length(hlinksInds) > 0){
    
      hlinks <- hlinks[hlinksInds]
      for(i in 1:length(hlinksInds)){
        targets <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
        targets <- replaceXMLEntities(gsub('"$', "", targets))
        
        hSheet <- sheetNumber[hlinksInds[i]]
        names(wb$hyperlinks[[hSheet]]) <- targets  
      }  
    }
    
    ## drawings
    draw <- lapply(xml, function(x) x[grepl("drawings/drawing", x)])
    drawInds <- which(sapply(draw, length) > 0)
    
    if(length(drawingRelsXML) > 0){

      dRels <- lapply(drawingRelsXML, readLines, warn = FALSE)
      dRels <- unlist(lapply(dRels, removeHeadTag))
      dRels <- gsub("<Relationships .*?>", "", dRels)
      dRels <- gsub("</Relationships>", "", dRels)
    }
    
    if(length(drawingsXML) > 0){
      dXML <- lapply(drawingsXML, readLines, warn = FALSE)  
      dXML <- unlist(lapply(dXML, removeHeadTag))
      dXML <- gsub("<xdr:wsDr .*?>", "", dXML)
      dXML <- gsub("</xdr:wsDr>", "", dXML)
      
      ## split at one/two cell Anchor
      dXML <- regmatches(dXML, gregexpr("<xdr:...CellAnchor.*?</xdr:...CellAnchor>", dXML))
    } 

    if(length(drawInds) > 0){
      
      draw <- draw[drawInds]
      for(i in 1:length(drawInds)){
        
        dSheet <- sheetNumber[drawInds[i]]
        
        ## dSheet has a relationship pointing to the drawing(i).xml file specified by target
        target <- unlist(lapply(draw[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
        target <- tail(strsplit(gsub('"$', "", target), split = "/")[[1]], 1)
        
        ## the media beloning to this sheet
        mediaInd <- grepl(target, drawingRelsXML)
        if(any(mediaInd))
          wb$drawings_rels[dSheet] <- dRels[mediaInd]
        
        drawingInd <- grepl(target, drawingsXML)
        if(any(drawingInd))
          wb$drawings[dSheet] <- dXML[drawingInd]
      
      }     
    }
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

  return(wb)
  
}

