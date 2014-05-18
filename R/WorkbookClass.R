


Workbook <- setRefClass("Workbook", fields = c(".rels",
                                               "app",
                                               "charts",
                                               "colWidths",
                                               "connections",
                                               "Content_Types",
                                               "core",
                                               "dataCount",
                                               "drawings",
                                               "drawings_rels",
                                               "externalLinks",
                                               "externalLinksRels",
                                               "freezePane",
                                               "headFoot",
                                               "hyperlinks",
                                               "media",
                                               "printerSettings",
                                               "queryTables",
                                               "rowHeights",
                                               "sharedStrings",
                                               "sheetData",
                                               "styleObjects",
                                               "styles",
                                               "tables",
                                               "tables.xml.rels",
                                               "theme",
                                               "workbook",
                                               "workbook.xml.rels",
                                               "worksheets",
                                               "worksheets_rels",
                                               "sheetOrder")
)


Workbook$methods(initialize = function(creator = Sys.info()[["login"]]){
  
  if(length(creator) == 0)
    creator <- ""
  
  Content_Types <<- genBaseContent_Type()
  .rels <<- genBaseRels()
  drawings <<- list()
  drawings_rels <<- list()
  media <<- list()
  charts <<- list()
  app <<- genBaseApp()
  core <<- genBaseCore(creator)
  workbook.xml.rels <<- genBaseWorkbook.xml.rels()
  theme <<- genBaseTheme()
  worksheets <<- list()
  worksheets_rels <<- list()
  sharedStrings <<- list()
  styles <<- genBaseStyleSheet()
  workbook <<- genBaseWorkbook()
  sheetData <<- list()
  rowHeights <<- list()
  styleObjects <<- list()
  colWidths <<- list()
  dataCount <<- list()
  freezePane <<- list()
  tables <<- NULL
  tables.xml.rels <<- NULL
  queryTables <<- NULL
  connections <<- NULL
  externalLinks <<- NULL
  externalLinksRels <<- NULL
  headFoot <<- data.frame("text" = rep(NA, 6), "pos" = c("left", "center", "right"), "head" = c("head", "head", "head", "foot", "foot", "foot"), stringsAsFactors = FALSE)
  printerSettings <<- list()
  hyperlinks <<- list()
  sheetOrder <<- NULL
  
  attr(sharedStrings, "uniqueCount") <<- 0
  
})

Workbook$methods(zipWorkbook = function(zipfile, files, flags = "-r1", extras = "", zip = Sys.getenv("R_ZIPCMD", "zip"), quiet = TRUE){ 
  
    ## code from utils::zip function (modified to not print)
    args <- c(flags, shQuote(path.expand(zipfile)), shQuote(files), extras)
    
    if(quiet){
      invisible(system2(zip, args, stdout = NULL))
    }else{
      if (.Platform$OS.type == "windows") 
        invisible(system2(zip, args, invisible = TRUE))
      else invisible(system2(zip, args))
    }
    
    invisible(0)
})


Workbook$methods(addWorksheet = function(sheetName, showGridLines = TRUE){
    
  newSheetIndex = length(worksheets) + 1
    
  ##  Add sheet to workbook.xml
  workbook$sheets <<- c(workbook$sheets, sprintf('<sheet name="%s" sheetId="%s" r:id="rId%s"/>', sheetName, newSheetIndex, newSheetIndex))
  
  ## append to worksheets list
  worksheets <<- append(worksheets, genBaseSheet(sheetName, showGridLines))
  
  ## update content_tyes
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex))
  
  ## Update xl/rels
  workbook.xml.rels <<- c(workbook.xml.rels,
                          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>', newSheetIndex)
                        )
  
  ## add a drawing.xml for the worksheet
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', newSheetIndex))
    
  ## create sheet.rels to simplify id assignment
  worksheets_rels[[newSheetIndex]] <<- genBaseSheetRels(newSheetIndex)
  drawings_rels[[newSheetIndex]] <<- ""
  drawings[[newSheetIndex]] <<- ""
  
  sheetData[[newSheetIndex]] <<- list()
  rowHeights[[newSheetIndex]] <<- list()
  colWidths[[newSheetIndex]] <<- list()
  freezePane[[newSheetIndex]] <<- list()
  printerSettings[[newSheetIndex]] <<- genPrinterSettings()
  hyperlinks[[newSheetIndex]] <<- ""
  
  dataCount[[newSheetIndex]] <<- 0
  sheetOrder <<- c(sheetOrder, newSheetIndex)
  
  invisible(newSheetIndex)
  
})

Workbook$methods(saveWorkbook = function(quiet = TRUE){
  
  ## temp directory to save XML files prior to compressing
  tmpDir <- file.path(tempfile(pattern="workbookTemp_"))
    
  if(file.exists(tmpDir))
    unlink(tmpDir, recursive = TRUE, force = TRUE)
  
  success <- dir.create(path = tmpDir, recursive = TRUE)
  if(!success)
    stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
    
  .self$preSaveCleanUp()
    
  nSheets <- length(worksheets)
  nThemes <- length(theme)

  relsDir <- file.path(tmpDir, "_rels")
  dir.create(path = relsDir, recursive = TRUE)
  
  docPropsDir <- file.path(tmpDir, "docProps")
  dir.create(path = docPropsDir, recursive = TRUE)
  
  xlDir <- file.path(tmpDir, "xl")
  dir.create(path = xlDir, recursive = TRUE)
  
  xlrelsDir <- file.path(tmpDir, "xl","_rels")
  dir.create(path = xlrelsDir, recursive = TRUE)
  
  xlTablesDir <- file.path(tmpDir, "xl","tables")
  dir.create(path = xlTablesDir, recursive = TRUE)
  
  xlTablesRelsDir <- file.path(xlTablesDir, "_rels")
  dir.create(path = xlTablesRelsDir, recursive = TRUE)
      
  if(length(media) > 0){
    xlmediaDir <- file.path(tmpDir, "xl", "media")
    dir.create(path = xlmediaDir, recursive = TRUE)
  }
  
  if(nThemes > 0){
    xlthemeDir <- file.path(tmpDir, "xl", "theme")
    dir.create(path = xlthemeDir, recursive = TRUE)
  }
    
  ## will always have drawings
  xlworksheetsDir <- file.path(tmpDir, "xl", "worksheets")
  dir.create(path = xlworksheetsDir, recursive = TRUE)
  
  xlworksheetsRelsDir <- file.path(tmpDir, "xl", "worksheets", "_rels")
  dir.create(path = xlworksheetsRelsDir, recursive = TRUE)
    
  xldrawingsDir <- file.path(tmpDir, "xl", "drawings")
  dir.create(path = xldrawingsDir, recursive = TRUE)
    
  xldrawingsRelsDir <- file.path(tmpDir, "xl", "drawings", "_rels")
  dir.create(path = xldrawingsRelsDir, recursive = TRUE)
  
  ##charts dir
  chartsDir <- file.path(tmpDir, "xl", "charts")
  dir.create(path = chartsDir, recursive = TRUE)
  
  printDir <- file.path(tmpDir, "xl", "printerSettings")
  dir.create(path = printDir, recursive = TRUE)
  
  
  ## Write content
    
  ## write .rels
  .Call("openxlsx_writeFile", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                            pxml(.rels),
                            '</Relationships>',
                            file.path(relsDir, ".rels"))
  
                      
  ## write app.xml
  .Call("openxlsx_writeFile", '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
                            pxml(app),
                            '</Properties>',
                            file.path(docPropsDir, "app.xml"))
  
  ## write core.xml
  .Call("openxlsx_writeFile", '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
                            pxml(core),
                            '</coreProperties>',
                            file.path(docPropsDir, "core.xml"))
  
  ## write workbook.xml.rels
  .Call("openxlsx_writeFile", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                            pxml(workbook.xml.rels),
                            '</Relationships>',
                            file.path(xlrelsDir, "workbook.xml.rels"))

  ## write tables
  if(length(unlist(tables)) > 0){
    for(i in 1:length(unlist(tables)))
      .Call("openxlsx_writeFile", '', pxml(unlist(tables)[[i]]), '', file.path(xlTablesDir, sprintf("table%s.xml", i+2)))
      .Call("openxlsx_writeFile", '', tables.xml.rels[[i]], '', file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", i+2)))
  }

    
  ## write query tables
  if(length(queryTables) > 0){
    xlqueryTablesDir <- file.path(tmpDir, "xl","queryTables")
    dir.create(path = xlqueryTablesDir, recursive = TRUE)
    
    for(i in 1:length(queryTables))
      .Call("openxlsx_writeFile", '', queryTables[[i]], '', file.path(xlqueryTablesDir, sprintf("queryTable%s.xml", i)))
  }
  
  ## connections
  if(length(connections) > 0)
      .Call("openxlsx_writeFile", '', connections, '', file.path(xlDir,"connections.xml"))
  
  ## externalLinks
  if(length(externalLinks)){
    externalLinksDir <- file.path(tmpDir, "xl","externalLinks")
    dir.create(path = externalLinksDir, recursive = TRUE)
    
    for(i in 1:length(externalLinks))
      .Call("openxlsx_writeFile", '', externalLinks[[i]], '', file.path(externalLinksDir, sprintf("externalLink%s.xml", i)))
  }

  ## externalLinks rels
  if(length(externalLinksRels)){
    externalLinksRelsDir <- file.path(tmpDir, "xl","externalLinks", "_rels")
    dir.create(path = externalLinksRelsDir, recursive = TRUE)
    
    for(i in 1:length(externalLinksRels))
      .Call("openxlsx_writeFile", '', externalLinksRels[[i]], '', file.path(externalLinksRelsDir, sprintf("externalLink%s.xml.rels", i)))
  } 
  
  # printerSettings
  for(i in 1:nSheets)
   writeLines(printerSettings[[i]], file.path(printDir, sprintf("printerSettings%s.bin", i)))
  
  ## media (copy file from origin to destination)
  for(x in media)
    file.copy(x, file.path(xlmediaDir, names(media)[which(media == x)]))
  
  ## charts
  for(x in charts)
    file.copy(x, file.path(chartsDir, names(charts)[which(charts == x)]))
  
  ## will always have a theme
    lapply(1:nThemes, function(i){
      con <- file(file.path(xlthemeDir, paste0("theme", i, ".xml")), open = "wb")
      writeBin(charToRaw(pxml(theme[[i]])), con)
      close(con)        
    })
  
  ## write worksheet, worksheet_rels, drawings, drawing_rels
  .self$writeSheetDataXML(xldrawingsDir, xldrawingsRelsDir, xlworksheetsDir, xlworksheetsRelsDir)
        
  ## write shareStrings.xml
  ct <- Content_Types
  if(length(sharedStrings) > 0){
    .Call("openxlsx_writeFile",
          sprintf('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%s" uniqueCount="%s">',   length(sharedStrings),   attr(sharedStrings, "uniqueCount")),
          paste(sharedStrings, collapse = ""),
          "</sst>",
          file.path(xlDir,"sharedStrings.xml"))
  }else{
    
    ## Remove relationship to sharedStrings
    ct <- ct[!grepl("sharedStrings", ct)]
  }

  

  ## write [Content_type]       
  .Call("openxlsx_writeFile", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        pxml(ct),
        '</Types>', 
        file.path(tmpDir, "[Content_Types].xml"))
  
  
  styleXML <- styles
  styleXML$numFmts <- paste0(sprintf('<numFmts count="%s">', length(styles$numFmts)), pxml(styles$numFmts), '</numFmts>')
  styleXML$fonts <- paste0(sprintf('<fonts count="%s">', length(styles$fonts)), pxml(styles$fonts), '</fonts>')
  styleXML$fills <- paste0(sprintf('<fills count="%s">', length(styles$fills)), pxml(styles$fills), '</fills>')
  styleXML$borders <- paste0(sprintf('<borders count="%s">', length(styles$borders)), pxml(styles$borders), '</borders>')
  styleXML$cellStyleXfs <- c(sprintf('<cellStyleXfs count="%s">', length(styles$cellStyleXfs)), pxml(styles$cellStyleXfs), '</cellStyleXfs>')
  styleXML$cellXfs <- paste0(sprintf('<cellXfs count="%s">', length(styles$cellXfs)), pxml(styles$cellXfs), '</cellXfs>')
  styleXML$cellStyles <- paste0(sprintf('<cellStyles count="%s">', length(styles$cellStyles)), pxml(styles$cellStyles), '</cellStyles>')
  styleXML$dxfs <- ifelse(length(styles$dxfs) == 0, '<dxfs count="0"/>', paste0(sprintf('<dxfs count="%s">', length(styles$dxfs)), pxml(styles$dxfs), '</dxfs>'))
  
  ## write styles.xml
  .Call("openxlsx_writeFile", '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">',
        pxml(styleXML),
        '</styleSheet>',
        file.path(xlDir,"styles.xml"))
  
  ## write workbook.xml
  workbookXML <- workbook
  workbookXML$sheets <- paste0("<sheets>", pxml(workbookXML$sheets), "</sheets>")
  .Call("openxlsx_writeFile", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                            pxml(workbookXML),
                            '</workbook>',   
                            file.path(xlDir,"workbook.xml"))
  
  ## compress to xlsx
  setwd(tmpDir)
  zipWorkbook("temp.xlsx", list.files(tmpDir, recursive = TRUE, include.dirs = TRUE, all.files=TRUE), quiet = quiet)

  ## reset styles
  baseFont <- styles$fonts[[1]]
  styles <<- genBaseStyleSheet()
  styles$fonts[[1]] <<- baseFont

  invisible(tmpDir)
  
})



Workbook$methods(updateSharedStrings = function(uNewStr){
  
  ## Function will return named list of references to new strings  
  uStr <- uNewStr[which(!uNewStr %in% sharedStrings)]
  uCount <- attr(sharedStrings, "uniqueCount")
  sharedStrings <<- append(sharedStrings, uStr)
    
  attr(sharedStrings, "uniqueCount") <<- uCount + length(uStr)
  
  })
  
  

Workbook$methods(validateSheet = function(sheetName){
  
  exSheets <- names(worksheets)
  
  if(!is.numeric(sheetName))
    sheetName <- replaceIllegalCharacters(sheetName)
  
  if(is.null(exSheets))
    stop("Workbook does not contain any worksheets.", call.=FALSE)
  
  if(is.numeric(sheetName)){
    if(sheetName > length(exSheets))
      stop(sprintf("This Workbook only has %s sheets.", length(exSheets)), call.=FALSE)
    
  return(sheetName)
  
  }else if(!sheetName %in% exSheets){
    stop(sprintf("Sheet '%s' does not exist.", sheetName), call.=FALSE)
  }
  
  return(which(exSheets == sheetName))
  
})



Workbook$methods(getSheetName = function(sheetIndex){
    
  sheetNames <- names(worksheets)
  
  if(any(length(sheetNames) < sheetIndex))
    stop(sprintf("Sheet only contains %s sheet(s).", length(sheetNames)))

  sheetNames[sheetIndex]
  
})



Workbook$methods(buildTable = function(sheet, colNames, ref, showColNames, tableStyle){
  
  ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
  id <- as.character(length(tables) + 3)
  sheet = validateSheet(sheet)

  ## build table XML and save to tables field
  tables <<- c(tables, .Call("openxlsx_buildTableXML", id, ref, colNames, showColNames, tableStyle, PACKAGE = "openxlsx"))
  
  worksheets[[sheet]]$tableParts <<- append(worksheets[[sheet]]$tableParts, sprintf('<tablePart r:id="rId%s"/>', id))
    
  ## update Content_Types
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', id))
    
  ## create a table.xml.rels
  tables.xml.rels <<- append(tables.xml.rels, "")
  
  ## update worksheets_rels
  worksheets_rels[[sheet]] <<- c(worksheets_rels[[sheet]],
    sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',  id,  id))


})



Workbook$methods(writeData = function(df, sheet, startRow, startCol, colNames){
  
  sheet = validateSheet(sheet)
  nCols <- ncol(df)
  nRows <- nrow(df)  
  
  colClasses <- sapply(df, function(x) class(x)[[1]])
  
  ## convert any Dates to integers and create date style object
  if(any(c("Date", "POSIXct", "POSIXt") %in% colClasses)){
    dInds <- which(sapply(colClasses, function(x) "Date" %in% x))
    for(i in dInds)
      df[,i] <- as.integer(df[,i]) + 25569
    
    pInds <- which(sapply(colClasses, function(x) any(c("POSIXct", "POSIXt") %in% x)))
    for(i in pInds)
      df[,i] <- as.integer(df[,i])/86400 + 25569
  }
  
  ## convert any Dates to integers and create date style object
  if(any(c("currency", "accounting", "percentage", "3") %in% tolower(colClasses))){
    cInds <- which(sapply(colClasses, function(x) any(c("accounting", "currency", "percentage", "3") %in% tolower(x))))
    for(i in cInds)
      df[,i] <- as.numeric(gsub("[^0-9\\.]", "", df[,i]))
  }
  
  
  colClasses <- sapply(df, function(x) class(x)[[1]])
  
  t <- .Call("openxlsx_buildCellTypes", colClasses, nRows, PACKAGE = "openxlsx")
  
  if("logical" %in% colClasses){
    df[df == TRUE] <- "1"
    df[df == FALSE] <- "0"
  }
  
  v <- as.character(t(as.matrix(df)))
  v[is.na(v)] <- as.character(NA)
  t[is.na(v)] <- as.character(NA)
  
  #prepend column headers 
  if(colNames){
    t <- c(rep.int('s', nCols), t)
    v <- c(names(df), v)
    nRows <- nRows + 1
  }
  
  ## create references
  r <- .Call("openxlsx_ExcelConvertExpand", startCol:(startCol+nCols-1), LETTERS, as.character(startRow:(startRow+nRows-1)))

  
  ##Append hyperlinks, convert h to s in cell type
  if("hyperlink" %in% tolower(colClasses)){
    
    hInds <- which(t == "h")
    t[hInds] <- "s"
    
    exHlinks <- hyperlinks[[sheet]]
    exhlinkRefs <- names(hyperlinks[[sheet]])
    
    newHlinks <- r[hInds]
    names(newHlinks) <- v[hInds]
    
    if(exHlinks[[1]] == ""){
      hyperlinks[[sheet]] <<- newHlinks
    }else{
      allHlinks <- c(exHlinks, newHlinks)
      allHlinks <- allHlinks[!duplicated(allHlinks, fromLast = TRUE)]
      allHlinks <- allHlinks[order(nchar(allHlinks), allHlinks)]
      hyperlinks[[sheet]] <<- allHlinks
    }
    
  }
  
  ## convert all strings to references in sharedStrings and update values (v)
  strFlag <- which(t == "s")
  newStrs <- v[strFlag]
  if(length(newStrs) > 0){
  
    newStrs <- replaceIllegalCharacters(newStrs)
    newStrs <- paste0("<si><t>", newStrs, "</t></si>")
    
    uNewStr <- unique(newStrs)
    
    .self$updateSharedStrings(uNewStr)  
    v[strFlag] <- as.integer(match(newStrs, sharedStrings) - 1)
  }
  

  
 
  ## Create cell list of lists

  cells <- .Call("openxlsx_buildCellList", r , t ,v , PACKAGE="openxlsx")
  names(cells) <- as.integer(names(r))

  if(dataCount[[sheet]] > 0){
    sheetData[[sheet]] <<- .Call("openxlsx_uniqueCellAppend", sheetData[[sheet]], r, cells, PACKAGE = "openxlsx")
    if(attr(sheetData[[sheet]], "overwrite"))
      warning("Overwriting existing cell data.", call. = FALSE)
    
    attr(sheetData[[sheet]], "overwrite") <<- NULL
  }else{
    sheetData[[sheet]] <<- cells
  }
  
  dataCount[[sheet]] <<- dataCount[[sheet]] + 1
  
})          



  
  

Workbook$methods(updateCellStyles = function(sheet, rows, cols, styleId){
    
  if(length(rows) == 0)
    return(NULL)

  ## convert sheet name to index
  sheet <- which(names(worksheets) == sheet)
  sheetData[[sheet]] <<- .Call("openxlsx_writeCellStyles", sheetData[[sheet]], as.character(rows), cols, as.character(styleId), LETTERS)

})


Workbook$methods(updateStyles = function(style){
    
  
  ## Updates styles.xml
  numFmtId <- 0
  fontId <- 0
  fillId <- 0
  borderId <- 0
  xfId <- 0
  
  xfNode <- list(numFmtId = 0,
                 fontId = 0,
                 fillId = 0,
                 borderId = 0,
                 xfId = 0)

  
  alignmentFlag <- FALSE
  
  ## Font
  if(!is.null(style$fontName) | !is.null(style$fontSize) |  !is.null(style$fontColour)){
  
    fontNode <- .self$createFontNode(style)
    fontId <- which(styles$font == fontNode)-1
      
    if(length(fontId) == 0){
            
        fontId <- length(styles$fonts)
        styles$fonts <<- append(styles[["fonts"]], fontNode)        
        
    }
      
    xfNode$fontId <- fontId
    xfNode <- append(xfNode, list("applyFont" = 1))
  }
  

  ## numFmt
  if(as.numeric(style$numFmt$numFmtId) > 0){
    numFmtId <- style$numFmt$numFmtId
    if(as.numeric(numFmtId) > 163){
      
      tmp <- style$numFmt$formatCode
      
      styles$numFmts <<- unique(c(styles$numFmts,
                              sprintf('<numFmt numFmtId="%s" formatCode="%s"/>', numFmtId, tmp)
                             ))
    }
      
    xfNode$numFmtId <- numFmtId
    xfNode <- append(xfNode, list("applyNumberFormat" = 1))
    
  }
  
  ## Fill
  if(!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg)){
    
    fillNode <- .self$createFillNode(style$fill)
    fillId <- which(styles$fill == fillNode)-1
    
    if(length(fillId) == 0){      
      fillId <- length(styles$fills)
      styles$fills <<- c(styles$fills, fillNode)
    }
    xfNode$fillId <- fillId
    xfNode <- append(xfNode, list("applyFill" = 1))
  }
 
  ## Border
  if(!all(is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom)))){

    borderNode <- .self$createBorderNode(style)
    borderId <- which(styles$borders == borderNode)-1

    if(length(borderId) == 0){
      borderId <- length(styles$borders)
      styles$borders <<- c(styles$borders, borderNode)
    }
        
    xfNode$borderId <- borderId
    xfNode <- append(xfNode, list("applyBorder" = 1))
  }
  
  ## Alignment
  if(!is.null(style$halign) | !is.null(style$valign) | !is.null(style$wrapText) | !is.null(style$textRotation)){
        
    attrs <- list()
    alignNode <- "<alignment"
    
    if(!is.null(style$textRotation))
      alignNode <- paste(alignNode, sprintf('textRotation="%s"', style$textRotation))
    
    if(!is.null(style$halign))
      alignNode <- paste(alignNode, sprintf('horizontal="%s"', style$halign))
    
    if(!is.null(style$valign))
      alignNode <- paste(alignNode, sprintf('vertical="%s"', style$valign))
    
    if(style$wrapText)
      alignNode <- paste(alignNode, 'wrapText="1"')
    
    alignNode <- paste0(alignNode, "/>")
    
    alignmentFlag <- TRUE
    xfNode <- append(xfNode, list("applyAlignment" = 1))
  }
    
  styleId <- length(styles$cellXfs)
  
  if(alignmentFlag){
    xfNode <- paste0("<xf ", paste(paste0(names(xfNode), '="',xfNode, '"'), collapse = " "), ">", alignNode, '</xf>')  
  }else{
    xfNode <- paste0("<xfv", paste(paste0(names(xfNode), '="',xfNode, '"'), collapse = " "), "/>")
  }
  
  styleId <- which(styles$cellXfs == xfNode) - 1
  if(length(styleId) == 0){
    styleId <- length(styles$cellXfs)
    styles$cellXfs <<- c(styles$cellXfs, xfNode)
  }
  
  
  return(styleId)

})


Workbook$methods(getBaseFont = function(){
  
  baseFont <- styles$fonts[[1]]
  
  sz <- getAttrs(baseFont, "<sz ")
  colour <- getAttrs(baseFont, "<color ")
  name <- getAttrs(baseFont, "<name ")
  family <- getAttrs(baseFont, "<family ")
  
  list("size" = sz,
    "colour" = colour,
    "name" = name)
  
})




Workbook$methods(createFontNode = function(style){
  
  baseFont <- .self$getBaseFont()
  
  fontNode <- "<font>"
  
  ## size
  if(is.null(style$fontSize[[1]])){
    fontNode <- paste0(fontNode, sprintf('<sz %s="%s"/>', names(baseFont$size), baseFont$size))
  }else{
    fontNode <- paste0(fontNode, sprintf('<sz %s="%s"/>', names(style$fontSize), style$fontSize))
  }
  
  ## colour
  if(is.null(style$fontColour[[1]])){
    fontNode <- paste0(fontNode, sprintf('<color %s="%s"/>', names(baseFont$colour), baseFont$colour))
  }else{
    fontNode <- paste0(fontNode, sprintf('<color %s="%s"/>', names(style$fontColour), style$fontColour))
  }
  
  
  ## name
  if(is.null(style$fontName[[1]])){
    fontNode <- paste0(fontNode, sprintf('<name %s="%s"/>', names(baseFont$name), baseFont$name))
  }else{
    fontNode <- paste0(fontNode, sprintf('<name %s="%s"/>', names(style$fontName), style$fontName))
  }
  
  ### Create new font and return Id  
  if(!is.null(style$fontFamily))
    fontNode <- paste0(fontNode, sprintf('<family val = "%s"/>', style$fontFamily))
  
  if(!is.null(style$fontScheme))
    fontNode <- paste0(fontNode, sprintf('<scheme val = "%s"/>', style$fontScheme))
  
  if("BOLD" %in% style$fontDecoration)
    fontNode <- paste0(fontNode, '<b/>')
  
  if("ITALIC" %in% style$fontDecoration)
    fontNode <- paste0(fontNode, '<i/>')
  
  if("UNDERLINE" %in% style$fontDecoration)
    fontNode <- paste0(fontNode, '<u val = "single"/>')
  
  if("UNDERLINE2" %in% style$fontDecoration)
    fontNode <- paste0(fontNode, '<u val = "double"/>')
  
  if("STRIKEOUT" %in% style$fontDecoration)
    fontNode <- paste0(fontNode, '<strike/>')
  
  paste0(fontNode, "</font>")
  
})
  

Workbook$methods(createBorderNode = function(style){
    
  borderNode <- "<border>"
  
  if(!is.null(style$borderLeft))
    borderNode <- paste0(borderNode, sprintf('<left style="%s">', style$borderLeft), sprintf('<color %s="%s"/>', names(style$borderLeftColour), style$borderLeftColour), '</left>')
  
  if(!is.null(style$borderRight))
    borderNode <- paste0(borderNode, sprintf('<right style="%s">', style$borderRight), sprintf('<color %s="%s"/>', names(style$borderRightColour), style$borderRightColour), '</right>')
  
  if(!is.null(style$borderTop))
    borderNode <- paste0(borderNode, sprintf('<top style="%s">', style$borderTop), sprintf('<color %s="%s"/>', names(style$borderTopColour), style$borderTopColour), '</top>')
  
  if(!is.null(style$borderBottom))
    borderNode <- paste0(borderNode, sprintf('<bottom style="%s">', style$borderBottom), sprintf('<color %s="%s"/>', names(style$borderBottomColour), style$borderBottomColour), '</bottom>')
  
  paste0(borderNode, "</border>")
  
})


Workbook$methods(createFillNode = function(fill, patternType="solid"){
      
  fillNode <- '<fill>'
  fillNode <- paste0(fillNode, sprintf('<patternFill patternType="%s">', patternType))
  
  
  
  
  if(!is.null(fill$fillFg))
    fillNode <- paste0(fillNode, sprintf('<fgColor %s/>', paste(paste0(names(fill$fillFg), '="', fill$fillFg, '"'), collapse = " ")))
  
  if(!is.null(fill$fillBg))
    fillNode <- paste0(fillNode, sprintf('<bgColor %s/>', paste(paste0(names(fill$fillBg), '="', fill$fillBg, '"'), collapse = " ")))
  
  fillNode <- paste0(fillNode, "</patternFill>")
  fillNode <- paste0(fillNode, "</fill>")
  
  return(fillNode)
  
})




    


Workbook$methods(setSheetName = function(sheet, newSheetName){
  
  if(newSheetName %in% names(worksheets))
    stop(sprintf("Sheet %s already exists!", newSheetName))  
  
  sheet <- validateSheet(sheet)
  
  oldName <- names(worksheets)[[sheet]]
  names(worksheets)[[sheet]] <<- newSheetName
  
  ## Rename in workbook
  sheetId <- regmatches(workbook$sheets[[sheet]], regexpr('(?<=sheetId=")[0-9]+', workbook$sheets[[sheet]], perl = TRUE))
  workbook$sheets[[sheet]] <<- sprintf('<sheet name="%s" sheetId="1" r:id="rId%s"/>', newSheetName, sheetId)
  
  ## remove styleObjects
  if(length(styleObjects) > 0){
    styleObjects <<- lapply(styleObjects, function(x){
      x$cells <- lapply(x$cells, function(s){ if(s$sheet == oldName) s$sheet <- newSheetName; s})
      x
    })
  }
  
  
})


Workbook$methods(writeSheetDataXML = function(xldrawingsDir, xldrawingsRelsDir, xlworksheetsDir, xlworksheetsRelsDir){
    
  ## write worksheets  
  nSheets <- length(worksheets)
  header <- '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' 
  
  for(i in 1:nSheets){
    
    ## Write drawing i (will always exist) skip those that are empty
    if(any(drawings[[i]] != "")){
      .Call("openxlsx_writeFile", 
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            pxml(drawings[[i]]),
            '</xdr:wsDr>',
            file.path(xldrawingsDir, paste0("drawing", i,".xml")))
      
    .Call("openxlsx_writeFile", 
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          pxml(drawings_rels[[i]]),
          '</Relationships>',
          file.path(xldrawingsRelsDir, paste0("drawing", i,".xml.rels")))
      }else{
        worksheets[[i]]$drawing <<- NULL
      }
    
    ## Write worksheets
    ws <- worksheets[[i]]
    if(!is.null(ws$cols))
      ws$cols <- pxml(c("<cols>", worksheets[[i]]$cols, "</cols>"))
    
    if(length(freezePane[[i]]) > 0)
      ws$sheetViews <- paste0('<sheetViews><sheetView workbookViewId=\"0\" tabSelected=\"TRUE\">', freezePane[[i]], '</sheetView></sheetViews>')
    
    if(length(worksheets[[i]]$mergeCells) > 0)
      ws$mergeCells <- paste0(sprintf('<mergeCells count="%s">', length(worksheets[[i]]$mergeCells)), pxml(worksheets[[i]]$mergeCells), '</mergeCells>')
    
    if(length(worksheets[[i]]$tableParts) > 0)
      ws$tableParts <- paste0(sprintf('<tableParts count="%s">', length(worksheets[[i]]$tableParts)), pxml(worksheets[[i]]$tableParts), '</tableParts>')
       
    if(hyperlinks[[i]][[1]] != ""){
      nTables <- length(tables)
      nHLinks <- length(hyperlinks[[i]])
      hInds <- 1:nHLinks + 3 + nTables-1
      ws$hyperlinks <- paste0('<hyperlinks>', paste(sprintf('<hyperlink ref="%s" r:id="rId%s"/>', hyperlinks[[i]], hInds), collapse = ""), '</hyperlinks>')  
    }
  
    sheetDataInd <- which(names(ws) == "sheetData")
    prior <- paste0(header, pxml(ws[1:(sheetDataInd-1)]))
    post <- paste0(pxml(ws[(sheetDataInd+1):length(ws)]), "</worksheet>")
    

    if(dataCount[[i]] == 1 & length(rowHeights[[i]]) == 0){
            
      .Call("openxlsx_quickBuildCellXML",
            prior,
            post,
            sheetData[[i]],
            as.integer(names(sheetData[[i]])),
            file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)),
            PACKAGE = "openxlsx")
    
    }else{
      
      if(length(rowHeights[[i]]) > 0)
        rowHeights[[i]] <<- rowHeights[[i]][order(as.numeric(names(rowHeights[[i]])))]
      
     .Call("openxlsx_buildCellXML",
            prior,
            post,
            sheetData[[i]],
            rowHeights[[i]],
            orderCellRef,
            file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)),
            PACKAGE="openxlsx")
    }
    
    ## write worksheet rels
    if(length(worksheets_rels[[i]]) > 0 ){
      
      ws_rels <- worksheets_rels[[i]]
      
      if(hyperlinks[[i]][[1]] != "")
        ws_rels <- c(ws_rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>', hInds, iconv(names(hyperlinks[[i]]), to = "UTF-8")))
          
      .Call("openxlsx_writeFile", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">', 
            pxml(ws_rels),
            '</Relationships>',
            file.path(xlworksheetsRelsDir, sprintf("sheet%s.xml.rels", i))
          )
    }
    
  } ## end of loop through 1:nSheets
       
  invisible(0)
  
})



Workbook$methods(setColWidths = function(sheet){
  
  sheet = validateSheet(sheet)
    
  widths <- unlist(lapply(colWidths[[sheet]], "[[", "width"))
  cols <- unlist(lapply(colWidths[[sheet]], "[[", "col"))
  
  autoColsInds <- which(widths == "auto")
  autoCols <- cols[autoColsInds]
  if(any(widths != "auto"))
    widths[widths != "auto"] <- as.numeric(widths[widths != "auto"]) + 0.71
  
  
  if(length(autoCols) > 0){
    
    ## get all cell values
    vals <- unlist(lapply(sheetData[[sheet]], "[[", "v"))
    colInds <- unlist(lapply(sheetData[[sheet]], "[[", "r"))
    
    stringInds <- which(unlist(lapply(sheetData[[sheet]], "[[", "t")) == "s")
    numericInds <- unlist(lapply(sheetData[[sheet]], "[[", "t")) == "n"
    
    ## replace string values and remove sharedString tags
    vals[stringInds] <- sharedStrings[(as.numeric(vals[stringInds]) + 1)]
    
    ## charLengths
    charLengths <- nchar(vals)
    
    ## truncate long numerics to 11 characters
    charLengths[charLengths > 11 & numericInds] <- 11
    names(charLengths) <- .Call("openxlsx_RcppConvertFromExcelRef", gsub("[0-9]+", "", colInds))
    
    
#     ##get Cell Styles
#     styleInd <- unlist(lapply(sheetData[[sheet]], "[[", "s"))
#     size <- list()
#     size[which(styleInd == NULL)] <- 11 
    
    calcWidths <- round(floor((charLengths*9 + 5)*256 / 9) / 256, 4) + 0.8 ## 0.8 My own adjustment 
    calcWidths <- calcWidths[!is.na(names(calcWidths))]
    
    
    calcWidths <- lapply(unique(names(calcWidths)), function(x) calcWidths[names(calcWidths) == x])
    
    
    maxWidths <- lapply(calcWidths, max)
    maxWidths[maxWidths < 8.43] <- 9.15 ## really 8.43 (For some reason excel subtracts 0.72)
    maxWidths[maxWidths > 50] <- 50.72    

    widths[autoColsInds] <- maxWidths[autoColsInds]
  }
    
  ## Calculate width of auto
  colNodes <- sprintf('<col min="%s" max="%s" width="%s" customWidth="1"/>', cols, cols, widths)
   
  ## Remove any existing widths that appear in cols
  flag <- !worksheets[[sheet]]$cols %in% cols
  if(any(flag))
    worksheets[[sheet]]$cols <<- worksheets[[sheet]]$cols[flag]

  ## Append new col widths XML to worksheets[[sheet]]$cols
  worksheets[[sheet]]$cols <<- append(worksheets[[sheet]]$cols, colNodes)
  
})


Workbook$methods(setRowHeights = function(sheet, rows, heights){
  
  sheet = validateSheet(sheet)
  
  ## remove any conflicting heights
  flag <- names(rowHeights[[sheet]]) %in% rows
  if(any(flag))
     rowHeights[[sheet]] <<- rowHeights[[sheet]][!flag]
  
  allRowHeights <- c(rowHeights[[sheet]], heights)
  allRowHeights <- allRowHeights[order(names(allRowHeights))]
  
  rowHeights[[sheet]] <<- allRowHeights
  
})


Workbook$methods(deleteWorksheet = function(sheet){
  
  # To delete a worksheet
  # Remove colwidths element
  # Remove drawing partname from Content_Types (drawing(sheet).xml)
  # Remove highest sheet from Content_Types
  # Remove drawings element
  # Remove drawings_rels element
  # Remove rowheights element
  # Remove sheetData
  # Remove styleObjects on sheet
  # Remove last sheet element from workbook
  # Remove last sheet element from workbook.xml.rels
  # Remove element from worksheets
  # Remove element from worksheets_rels
  # Remove Freeze Pane
  # Remove dataCount
  # Remove hyperlinks
  # Reduce calcChain i attributes & remove calcs on sheet
  # Remove sheet from sheetOrder

  sheet <- validateSheet(sheet)
  sheetNames <- names(worksheets)
  nSheets <- length(unlist(sheetNames))
  sheetName <- sheetNames[[sheet]]
      
  colWidths[[sheet]] <<- NULL
  
  ## remove last drawings(sheet).xml from Content_Types
  Content_Types <<- Content_Types[!grepl(sprintf("drawing%s.xml", nSheets), Content_Types)]
    
  ## remove highest sheet
  Content_Types <<- Content_Types[!grepl(sprintf("sheet%s.xml", nSheets), Content_Types)]
  
  drawings[[sheet]] <<- NULL
  drawings_rels[[sheet]] <<- NULL
  rowHeights[[sheet]] <<- NULL
  sheetData[[sheet]] <<- NULL
  hyperlinks[[sheet]] <<- NULL
  sheetOrder <<- c(sheetOrder[sheetOrder < sheet], sheetOrder[sheetOrder > sheet] - 1)
    
  ## remove styleObjects
  if(length(styleObjects) > 0){
    styleObjects <<- lapply(styleObjects, function(x){
      flag <- sapply(x$cells, function(s) s$sheet == sheetName)
      if(all(flag))
        return(NULL)
      x$cells <- x$cells[!flag]
      x
    })
    
    ## Now remove NULL values
    styleObjects <<- styleObjects[!sapply(styleObjects, is.null)]
  }
  
  ## wont't remove tables and then won't need to reassign table r:id's
  worksheets[[sheet]] <<- NULL
  worksheets_rels[[sheet]] <<- NULL

  ## drawing will always be the first relationship and printerSettings second
  for(i in 1:(nSheets-1))
    worksheets_rels[[i]][1:2] <<- genBaseSheetRels(i)
  
  
  workbook$sheets <<- workbook$sheets[!grepl(sprintf('name="%s"', sheetName), workbook$sheets)]
  ## Reset rIds
  for(i in (sheet+1):nSheets)
    workbook$sheets <<- gsub(paste0("rId", i), paste0("rId", i-1), workbook$sheets)
  
  ## Can remove highest sheet
  workbook.xml.rels <<- workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), workbook.xml.rels)]
  freezePane[[sheet]] <<- NULL
  dataCount[[sheet]] <<- NULL

  ## calcChain
#   if(length(calcChain) > 0){
#     ## remove elements with i = sheet
#     pattern <- sprintf(' i="%s"', sheet)
#     calcChain <<- calcChain[!grepl(pattern, calcChain)]
#     
#     ## decrement calcChain i attributes
#     original <- (1:nSheets)[1:nSheets > sheet]
#     if(length(original) > 0){
#       replacements <- paste0(' i="', original-1, '"')
#       original <- paste0(' i="', original, '"')
#       for(i in 1:length(original))
#         calcChain <<- gsub(original[[i]], replacements[[i]], calcChain)
#     
#     }
#   }
  
  
  
})


Workbook$methods(addDXFS = function(style){
     
  dxf <- '<dxf>'
  dxf <- paste0(dxf, createFontNode(style))
  fillNode <- NULL
  
  if(!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg))
    dxf <- paste0(dxf, createFillNode(style$fill))
  
  if(any(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom))
    dxf <- paste0(dxf, createBorderNode(style))
  
  dxfId <- length(styles$dxfs)
  styles$dxfs <<- c(styles$dxfs, paste(dxf, "</dxf>"))
  
  return(dxfId)
})



Workbook$methods(conditionalFormatCell = function(sheet, startRow, endRow, startCol, endCol, dxfId, operator, value){
    
  sheet = validateSheet(sheet)
  sqref <- paste(getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol))), collapse = ":")
  
  ## Increment priority
  if(length((worksheets[[sheet]]$conditionalFormatting)) > 0){
    for(i in length(worksheets[[sheet]]$conditionalFormatting):1)
      worksheets[[sheet]]$conditionalFormatting[[i]] <<- gsub('(?<=priority=")[0-9]+', i+1, worksheets[[sheet]]$conditionalFormatting[[i]], perl = TRUE)
  }
  worksheets[[sheet]]$conditionalFormatting <<- append(worksheets[[sheet]]$conditionalFormatting,               
    sprintf('<conditionalFormatting sqref="%s"><cfRule type="cellIs" dxfId="%s" priority="1" operator="%s"><formula>%s</formula></cfRule></conditionalFormatting>', sqref, dxfId, operator, value)                
  )
  
  invisible(return(0))
  
})




Workbook$methods(mergeCells = function(sheet, startRow, endRow, startCol, endCol) {
  
  sqref <- getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol)))
  exMerges <- regmatches(worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))
  
  if(!is.null(exMerges)){
   
    comps <- lapply(exMerges, function(rectCoords) unlist(strsplit(rectCoords, split = ":")))  
    exMergedCells <- .Call("openxlsx_buildCellMerges", comps, PACKAGE = "openxlsx")
    newMerge <- unlist(.Call("openxlsx_buildCellMerges", list(sqref), PACKAGE = "openxlsx"))
    
   ## Error if merge intersects    
    mergeIntersections <- sapply(exMergedCells, function(x) any(x %in% newMerge))
    if(any(mergeIntersections))
      stop(sprintf("Merge intersects with exisiting merged cells: \n\t\t%s.\nRemove exiting merge first.", paste(exMerges[mergeIntersections], collapse = "\n\t\t")))
    
  }
  
  worksheets[[sheet]]$mergeCells <<- c(worksheets[[sheet]]$mergeCells, sprintf('<mergeCell ref="%s"/>', paste(sqref, collapse = ":")))
  
  
})



Workbook$methods(removeCellMerge = function(sheet, startRow, endRow, startCol, endCol){
  
  sheet = validateSheet(sheet)
  
  sqref <- getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol)))
  exMerges <- regmatches(worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))
  
  if(!is.null(exMerges)){
    
    comps <- lapply(exMerges, function(x) unlist(strsplit(x, split = ":")))  
    exMergedCells <- .Call("openxlsx_buildCellMerges", comps)
    newMerge <- unlist(.Call("openxlsx_buildCellMerges", list(sqref)))
    
    ## Error if merge intersects    
    mergeIntersections <- sapply(exMergedCells, function(x) any(x %in% newMerge))
    
  }
  
  ## Remove children
  worksheets[[sheet]]$mergeCells <<- worksheets[[sheet]]$mergeCells[!mergeIntersections]
    
  
})


  
  

Workbook$methods(freezePanes = function(sheet, firstActiveRow = NULL, firstActiveCol = NULL, firstRow = FALSE, firstCol = FALSE){
  
  sheet = validateSheet(sheet)
  paneNode <- NULL
  
    if(firstRow){
      paneNode <- '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
    }else if(firstCol){
      paneNode <- '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>'
    }
  
  if(is.null(paneNode)){
    
    attrs <- NULL
    
    if(firstActiveRow > 1)
      attrs <- c(attrs, sprintf('ySplit="%s"', firstActiveRow - 1))
    
    if(firstActiveCol > 1)
      attrs <- c(attrs, sprintf('xSplit="%s"', firstActiveCol - 1))
    
    if(firstActiveRow == 1){
      activePane <- "topRight"
    }else{
      activePane <- "bottomRight"
    }
    
    topLeftCell <- getCellRefs(data.frame(firstActiveRow, firstActiveCol))
    
    paneNode <- sprintf('<pane %s topLeftCell="%s" activePane="%s" state="frozen"/><selection pane="%s"/>', 
                    paste(attrs, collapse = " "), topLeftCell, activePane, activePane)
    
    } 
  
  freezePane[[sheet]] <<- paneNode
  
})



Workbook$methods(insertImage = function(sheet, file, startRow, startCol, width, height, rowOffset = 0, colOffset = 0){
  
  ## within the sheet the drawing node's Id refernce an id in the sheetRels
  ## sheet rels reference the drawingi.xml file
  ## drawingi.xml refernece drawingRels
  ## drawing rels reference an image in the media folder
  ## worksheetRels(sheet(i)) references drawings(j)
  
  sheet = validateSheet(sheet)
  
  imageType <- regmatches(file, gregexpr("\\.[a-zA-Z]*$", file))
  imageType <- gsub("^\\.", "", imageType)
  
  imageNo <- length((drawings[[sheet]])) + 1
  mediaNo <- length(media) + 1
  
  startCol <- convertFromExcelRef(startCol)

  ## update Content_Types
  if(!any(grepl(paste0("image/", imageType), Content_Types)))
    Content_Types <<- unique(c(sprintf('<Default Extension="%s" ContentType="image/%s"/>', imageType, imageType), Content_Types))
      
  ## drawings rels (Reference from drawings.xml to image file in media folder)
  drawings_rels[[sheet]] <<- c(drawings_rels[[sheet]], 
   sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%s.%s"/>', imageNo, mediaNo, imageType))
                
  ## write file path to media slot to copy across on save
  tmp <- file
  names(tmp) <- paste0("image", mediaNo,".", imageType)
  media <<- append(media, tmp)
            
  ## create drawing.xml    
  anchor <- '<xdr:oneCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'

  from <- sprintf(
  '<xdr:from xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
    <xdr:col xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:col>
    <xdr:colOff xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:colOff>
    <xdr:row xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:row>
    <xdr:rowOff xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">%s</xdr:rowOff>
  </xdr:from>' , startCol-1, colOffset, startRow - 1, rowOffset)

  drawingsXML <- paste0(
      anchor,
      from,
      sprintf('<xdr:ext xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" cx="%s" cy="%s"/>', width, height),
      genBasePic(imageNo),
      '<xdr:clientData/>',
      '</xdr:oneCellAnchor>'
    )

  
  ## append to workbook drawing
  drawings[[sheet]] <<- c(drawings[[sheet]], drawingsXML) 
    
})



Workbook$methods(preSaveCleanUp = function(){
    

  ## Steps
  # Order workbook.xml.rels:
  #   sheets -> style -> theme -> sharedStrings -> tables -> calcChain
  # Assign workbook.xml.rels children rIds, 1:length(workbook.xml.rels)
  # Assign workbook$sheets rIds 1:nSheets
  #
  ## drawings will always be r:id1 on worksheet
  ## tables will always have r:id equal to table xml file number tables/table(i).xml
  
  ## Every worksheet has a drawingXML as r:id 1
  ## Every worksheet has a printerSettings as r:id 2
  ## Tables from r:id 3 to nTables+3 - 1
  ## HyperLinks from nTables+3 to nTables+3+nHyperLinks-1
  
  nSheets <- length(worksheets)
  nThemes <- length(theme)
  nExtRefs <- length(externalLinks)
  
  ## add a worksheet if none added
  if(nSheets == 0){
    warning("Workbook does not contain any worksheets. A worksheet will be added.")
    .self$addWorksheet("Sheet 1")
    nSheets <- 1  
  }

  ## Assign relationship Ids  
  nChildren <- 1:length(workbook.xml.rels)
  splitRels <- strsplit(workbook.xml.rels, split = " ")
  nms <- lapply(splitRels, function(x) gsub("(.*)=(.*)", "\\1", x[2:length(x)]))
  rAttrs <- lapply(splitRels, function(x) gsub("/>$", "", gsub('(.*)="(.*)"', "\\2", x[2:length(x)])))   
  rAttrs <- lapply(nChildren, function(i) {names(rAttrs[[i]]) <- nms[[i]]; return(rAttrs[[i]])})
  
  rIds <- paste0("rId", nChildren)
  targets <- sapply(rAttrs, "[[", "Target")
  
  ## Reorderinf of workbook.xml.rels
  sheetInds <- which(targets %in% sprintf("worksheets/sheet%s.xml", 1:nSheets))
  sheetNumbers <- as.numeric(gsub("[^0-9]", "", targets[sheetInds], perl = TRUE))
  sheetInds <- sheetInds[order(sheetNumbers)]  ## make sure sheets will be in order
  
  ## get index of each child element for ordering
  stylesInd <- which(targets == "styles.xml")
  themeInd <- which(targets %in% sprintf("theme/theme%s.xml", 1:nThemes))
  connectionsInd <- which(targets %in% "connections.xml")

  extRefInds <- which(targets %in% sprintf("externalLinks/externalLink%s.xml", 1:nExtRefs))
  sharedStringsInd <- which(targets == "sharedStrings.xml")
  tableInds <- which(grepl("table[0-9]+.xml", targets))
  
  ## Reorder children of workbook.xml.rels
  workbook.xml.rels <<- workbook.xml.rels[c(sheetInds, extRefInds, themeInd, connectionsInd, stylesInd, sharedStringsInd, tableInds)]
   
  ## Re assign rIds to children of workbook.xml.rels
  workbook.xml.rels <<- unlist(lapply(1:length(workbook.xml.rels), function(i) {
                              gsub('(?<=Relationship Id="rId)[0-9]+', i, workbook.xml.rels[[i]], perl = TRUE)
                            }))
  
  ## Reassign rId to workbook sheet elements, (order sheets by sheetId first)
  sId <- as.numeric(unlist(regmatches(workbook$sheets, gregexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE))))
  workbook$sheets <<- sapply(order(sId), function(i) gsub('(?<=id="rId)[0-9]+', i, workbook$sheets[[i]], perl = TRUE))
  workbook$sheets <<- sapply(1:nSheets, function(i) gsub('(?<=sheetId=")[0-9]+', i, workbook$sheets[[i]], perl = TRUE))
  if(!is.null(sheetOrder))
    workbook$sheets <<- workbook$sheets[sheetOrder]
  
  ## update workbook r:id to match reordered workbook.xml.rels externalLink element
  if(length(extRefInds) > 0){
    newInds <- 1:length(extRefInds) + length(sheetInds)
    workbook$externalReferences <<- paste0("<externalReferences>",
                                   paste0(sprintf('<externalReference r:id=\"rId%s\"/>', newInds), collapse = ""),
                                   "</externalReferences>")
  }

  ## styles
  for(x in styleObjects){
    if(length(x$cells) > 0){
      sId <- .self$updateStyles(x$style)
      for(r in x$cells)
          .self$updateCellStyles(sheet = r$sheet, rows = r$rows, cols = r$cols, sId)
    }
  }
  
  ## Header footer
  if(any(!is.na(headFoot$text)))
    .self$setHeaderFooter()
  
  ## Make sure all rowHeights have rows, if not append them!
  for(i in 1:length(worksheets)){
    
    if(length(rowHeights[[i]]) > 0){
      missingRows <- names(rowHeights[[i]])[!names(rowHeights[[i]]) %in% names(sheetData[[i]])]
      if(length(missingRows) > 0){
        n <- length(missingRows)
        newCells <- .Call("openxlsx_buildCellList", rep(as.character(NA), n) , rep(as.character(NA), n) , rep(as.character(NA), n), PACKAGE="openxlsx")
        names(newCells) <- missingRows
        sheetData[[i]] <<- append(sheetData[[i]], newCells)
      }
    }
    
    ## write colwidth XML   
    if(length(colWidths[[i]]) > 0)
      invisible(.self$setColWidths(i))
    
  }
  
})





Workbook$methods(setHeaderFooter = function(){
  
    headerPos <- headFoot$pos[!is.na(headFoot$text) & headFoot$head == "head"]
    
    if(length(headerPos) > 0){
      headNode <- "<oddHeader>"
      if("left" %in% headerPos)
        headNode <- paste0(headNode, "&amp;L",  headFoot$text[headFoot$pos == "left" & headFoot$head == "head"]) 
      
      if("center" %in% headerPos)
        headNode <- paste0(headNode, "&amp;C",  headFoot$text[headFoot$pos == "center" & headFoot$head == "head"]) 
      
      if("right" %in% headerPos)
        headNode <- paste0(headNode, "&amp;R",  headFoot$text[headFoot$pos == "right" & headFoot$head == "head"]) 
      headNode <- paste0(headNode, "</oddHeader>")
    }else{
      headNode <- NULL
    }
    
    
    footerPos <- headFoot$pos[!is.na(headFoot$text) & headFoot$head == "foot"]
    
    if(length(footerPos) > 0){
      footNode <- "<oddFooter>"
      if("left" %in% footerPos)
        footNode <- paste0(footNode, "&amp;L",  headFoot$text[headFoot$pos == "left" & headFoot$head == "foot"]) 
      
      if("center" %in% footerPos)
        footNode <- paste0(footNode, "&amp;C",  headFoot$text[headFoot$pos == "center" & headFoot$head == "foot"]) 
      
      if("right" %in% footerPos)
        footNode <- paste0(footNode, "&amp;R",  headFoot$text[headFoot$pos == "right" & headFoot$head == "foot"]) 
      footNode <- paste0(footNode, "</oddFooter>")
    }else{
      footNode <- NULL
    }
    
    worksheets[[1]]$headerFooter <<- paste0("<headerFooter>", headNode, footNode, "</headerFooter>")

})





Workbook$methods(show = function(){
  
  
  exSheets <- names(worksheets)
  nSheets <- length(exSheets)
  nImages <- length(media)
  nCharts <- length(charts)
  nStyles <- length(styleObjects)
  
  exSheets <- replaceXMLEntities(exSheets)
  showText <- "A Workbook object.\n"
  
  ## worksheets
  if(nSheets > 0){
    showText <- c(showText, "\nWorksheets:\n")
    
    sheetTxt <- lapply(1:nSheets, function(i){
      
      tmpTxt <- sprintf('Sheet %s: "%s"\n', i, exSheets[[i]])

      if(length(rowHeights[[i]]) > 0){
        
        tmpTxt <- append(tmpTxt, c("\n\tCustom row heights (row: height)\n\t", 
                        paste(sprintf("%s: %s", names(rowHeights[[i]]), round(as.numeric(rowHeights[[i]]), 2)), collapse = ", ")
                                   )
        )
      }
      
      
      if(length(colWidths[[i]]) > 0){
        cols <- lapply(colWidths[[i]], "[[", "col")
        widths <- lapply(colWidths[[i]], "[[", "width")
          
        widths[widths != "auto"] <- as.numeric(widths[widths != "auto"])
        tmpTxt <- append(tmpTxt, c("\n\tCustom column widths (column: width)\n\t ",
                         paste(sprintf("%s: %s", cols, substr(widths,1,5)), collapse = ", "))
        )
        tmpTxt <- c(tmpTxt, "\n") 
      }
      c(tmpTxt, "\n\n")
    })
      
    showText <- c(showText, sheetTxt, "\n")
    
  }else{
    showText <- c(showText, "\nWorksheets:\n", "No worksheets attached")
  }
  
  ## images
  if(nImages > 0)
    showText <- c(showText, "\nImages:\n", sprintf('Image %s: "%s"\n', 1:nImages, media))
  
  if(nCharts > 0)
    showText <- c(showText, "\nCharts:\n", sprintf('Chart %s: "%s"\n', 1:nImages, media))
  
  showText <- c(showText, sprintf("Worksheet write order: %s", paste(sheetOrder, collapse = ", ")))
  
  ## styles
#   if(nStyles > 0){
#     
#     showText <- c(showText, "\nWorkbook Styles: \n")
#     
#     styleText <- lapply(1:nStyles, function(i) {
#       
#       sheet <- paste(unlist(lapply(styleObjects[[i]]$cells, "[[", "sheet")), collapse = '", "')
#       styleTxt = invisible(styleObjects[[i]]$style$show(FALSE))
#       styleShow <- c(sprintf('Sheet(s): "%s" \n\n', sheet), styleTxt[-1])
#       
#       styleShow <- c(styleShow, "\n")
#       paste(styleShow, collapse = " ")
#       
#     })
#         
#     showText <- c(showText, unlist(styleText))
#     
#   }
  
  cat(unlist(showText))
  
})




