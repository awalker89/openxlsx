


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
                                               
                                               "pivotTables",
                                               "pivotTables.xml.rels",
                                               "pivotDefinitions",
                                               "pivotRecords",
                                               "pivotDefinitionsRels",
                                               
                                               "queryTables",
                                               "rowHeights",
                                               "sharedStrings",
                                               "sheetData",
                                               "styleObjects",
                                               "styles",
                                               "styleInds",
                                               "tables",
                                               "tables.xml.rels",
                                               "theme",
                                               
                                               "vbaProject",
                                               
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
  styleInds <<- list()
  
  colWidths <<- list()
  dataCount <<- list()
  freezePane <<- list()
  tables <<- NULL
  tables.xml.rels <<- NULL
  queryTables <<- NULL
  connections <<- NULL
  externalLinks <<- NULL
  externalLinksRels <<- NULL
  headFoot <<- NULL
  printerSettings <<- list()
  hyperlinks <<- list()
  sheetOrder <<- NULL
  
  pivotTables <<- NULL
  pivotTables.xml.rels <<- NULL
  pivotDefinitions <<- NULL
  pivotRecords <<- NULL
  pivotDefinitionsRels <<- NULL
  
  vbaProject <<- NULL
  
  attr(sharedStrings, "uniqueCount") <<- 0
  
})

Workbook$methods(zipWorkbook = function(zipfile, files, flags = "-r1", extras = "", zip = Sys.getenv("R_ZIPCMD", "zip"), quiet = TRUE){ 
  
  ## code from utils::zip function (modified to not print)
  args <- c(flags, shQuote(path.expand(zipfile)), shQuote(files), extras)
  
  if(quiet){
    
    res <- invisible(suppressWarnings(system2(zip, args, stdout = NULL)))
    
  }else{
    if (.Platform$OS.type == "windows"){
      res <- invisible(suppressWarnings(system2(zip, args, invisible = TRUE)))
    }else{
      res <- invisible(suppressWarnings(system2(zip, args)))
    }
  }
  
  if(res != 0){
    stop("zipping up workbook failed. Please make sure Rtools is installed or a zip application is available to R.
         Try installr::install.rtools() on Windows.", call. = FALSE)
  }
  
  invisible(res)
})


Workbook$methods(addWorksheet = function(sheetName, showGridLines = TRUE, tabColour = NULL, zoom = 100,
                                         oddHeader = NULL, oddFooter = NULL,
                                         evenHeader = NULL, evenFooter = NULL,
                                         firstHeader = NULL, firstFooter = NULL){
  
  newSheetIndex = length(worksheets) + 1L
  
  ##  Add sheet to workbook.xml
  workbook$sheets <<- c(workbook$sheets, sprintf('<sheet name="%s" sheetId="%s" r:id="rId%s"/>', sheetName, newSheetIndex, newSheetIndex))
  
  ## append to worksheets list
  worksheets <<- append(worksheets, genBaseSheet(sheetName = sheetName, showGridLines = showGridLines, 
                                                 tabSelected = newSheetIndex == 1, 
                                                 tabColour = tabColour, zoom = zoom,
                                                 oddHeader = oddHeader, oddFooter = oddFooter,
                                                 evenHeader = evenHeader, evenFooter = evenFooter,
                                                 firstHeader = firstHeader, firstFooter = firstFooter))
  
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
  styleInds[[newSheetIndex]] <<- list()
  
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
  nPivots <- length(pivotTables)
  
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
  
  if(nPivots > 0){
    
    pivotTablesDir <- file.path(tmpDir, "xl", "pivotTables")
    dir.create(path = pivotTablesDir, recursive = TRUE)
    
    pivotTablesRelsDir <- file.path(tmpDir, "xl", "pivotTables", "_rels")
    dir.create(path = pivotTablesRelsDir, recursive = TRUE)
    
    pivotCacheDir <- file.path(tmpDir, "xl", "pivotCache")
    dir.create(path = pivotCacheDir, recursive = TRUE)
    
    pivotCacheRelsDir <- file.path(tmpDir, "xl", "pivotCache", "_rels")
    dir.create(path = pivotCacheRelsDir, recursive = TRUE)
    
    for(i in 1:nPivots){
      .Call("openxlsx_writeFile", "", pivotTables[[i]], "", file.path(pivotTablesDir, sprintf("pivotTable%s.xml", i)))
      .Call("openxlsx_writeFile", "", pivotTables.xml.rels[[i]], "", file.path(pivotTablesRelsDir, sprintf("pivotTable%s.xml.rels", i)))   
      .Call("openxlsx_writeFile", "", pivotDefinitions[[i]], "", file.path(pivotCacheDir, sprintf("pivotCacheDefinition%s.xml", i)))
      .Call("openxlsx_writeFile", "", pivotRecords[[i]], "", file.path(pivotCacheDir, sprintf("pivotCacheRecords%s.xml", i)))
      .Call("openxlsx_writeFile", "", pivotDefinitionsRels[[i]], "", file.path(pivotCacheRelsDir, sprintf("pivotCacheDefinition%s.xml.rels", i)))   
    }
    
  }
  
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
  if(length(unlist(tables, use.names = FALSE)) > 0){
    for(i in 1:length(unlist(tables, use.names = FALSE))){
      .Call("openxlsx_writeFile", '', pxml(unlist(tables, use.names = FALSE)[[i]]), '', file.path(xlTablesDir, sprintf("table%s.xml", i+2)))
      
      if(tables.xml.rels[[i]] != "")
        .Call("openxlsx_writeFile", '', tables.xml.rels[[i]], '', file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", i+2)))
    }
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
  
  ## VBA Macro
  if(!is.null(vbaProject))
    file.copy(vbaProject, xlDir)
  
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
  
  workbook$sheets <<- workbook$sheets[order(sheetOrder)] ## Need to reset sheet order to allow multiple savings
  
  ## compress to xlsx
  setwd(tmpDir)
  tmpFile <- tempfile(tmpdir = tmpDir, fileext = ".xlsx")
  tmpFile <- basename(tmpFile)
  
  zipWorkbook(tmpFile, list.files(tmpDir, recursive = TRUE, include.dirs = TRUE, all.files=TRUE), quiet = quiet)
  
  ## reset styles - maintain any changes to base font
  baseFont <- styles$fonts[[1]]
  styles <<- genBaseStyleSheet(styles$dxfs)
  styles$fonts[[1]] <<- baseFont
  
  invisible(list("tmpDir" = tmpDir, "tmpFile" = tmpFile))
  
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



Workbook$methods(buildTable = function(sheet, colNames, ref, showColNames, tableStyle, tableName, withFilter){
  
  ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
  id <- as.character(length(tables) + 3L)
  sheet = validateSheet(sheet)
  
  ## build table XML and save to tables field
  table <- sprintf('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%s" name="%s" displayName="%s" ref="%s"', id, tableName, tableName, ref)
  
  nms <- names(tables)
  tSheets <- attr(tables, "sheet")
  tNames <- attr(tables, "tableName") 
  
  tables <<- c(tables, .Call("openxlsx_buildTableXML", table, ref, colNames, showColNames, tableStyle, withFilter, PACKAGE = "openxlsx"))
  names(tables) <<- c(nms, ref)
  attr(tables, "sheet") <<- c(tSheets, sheet)
  attr(tables, "tableName") <<- c(tNames, tableName)
  
  worksheets[[sheet]]$tableParts <<- append(worksheets[[sheet]]$tableParts, sprintf('<tablePart r:id="rId%s"/>', id))
  
  ## update Content_Types
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', id))
  
  ## create a table.xml.rels
  tables.xml.rels <<- append(tables.xml.rels, "")
  
  ## update worksheets_rels
  worksheets_rels[[sheet]] <<- c(worksheets_rels[[sheet]],
                                 sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',  id,  id))
  
  
})



Workbook$methods(writeData = function(df, sheet, startRow, startCol, colNames, colClasses, hlinkNames, keepNA){
  
  sheet <- validateSheet(sheet)
  nCols <- ncol(df)
  nRows <- nrow(df)  
  
  allColClasses <- unlist(colClasses)  
  
  ## convert any Dates to integers and create date style object
  if(any(c("date", "posixct", "posixt") %in% allColClasses)){
    dInds <- which(sapply(colClasses, function(x) "date" %in% x))
    for(i in dInds)
      df[,i] <- as.integer(df[,i]) + 25569
    
    t <- format(Sys.time(), "%z")
    offSet <- suppressWarnings(ifelse(substr(t,1,1) == "+", 1L, -1L) * (as.integer(substr(t,2,3)) + as.integer(substr(t,4,5)) / 60) / 24)
    
    if(is.na(offSet))
      offSet <- 0
    
    pInds <- which(sapply(colClasses, function(x) any(c("posixct", "posixt", "posixlt") %in% x)))
    for(i in pInds)
      df[,i] <- as.numeric(as.POSIXct(df[,i])) / 86400 + 25569 + offSet
  }
  
  ## convert any Dates to integers and create date style object
  if(any(c("currency", "accounting", "percentage", "3", "comma") %in% allColClasses)){
    cInds <- which(sapply(colClasses, function(x) any(c("accounting", "currency", "percentage", "3", "comma") %in% tolower(x))))
    for(i in cInds)
      df[,i] <- as.numeric(gsub("[^0-9\\.-]", "", df[,i]))
  }
  
  if("hyperlink" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "hyperlink" %in% x)))
      class(df[,i]) <- "hyperlink"
  }
  
  ## convert scientific
  if("scientific" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "scientific" %in% x)))
      class(df[,i]) <- "numeric"
  }
  
  colClasses <- sapply(df, function(x) tolower(class(x))[[1]]) ## by here all cols must have a single class only
  
  ## convert logicals (Excel stores logicals as 0 & 1)
  if("logical" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "logical" %in% x)))
      class(df[,i]) <- "numeric"
  }
  
  
  
  ## convert all numerics to character (this way preserves digits)
  if("numeric" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "numeric" %in% x)))
      class(df[,i]) <- "character"
  }
  
  ## cell types
  t <- .Call("openxlsx_buildCellTypes", colClasses, nRows, PACKAGE = "openxlsx")
  
  ## cell values
  v <- as.character(t(as.matrix(df)))
  
  if(keepNA){
    t[is.na(v)] <- "e"
    v[is.na(v)] <- "#N/A"
  }else{
    t[is.na(v)] <- as.character(NA)  
    v[is.na(v)] <- as.character(NA)
  }
  
  
  #prepend column headers 
  if(colNames){
    t <- c(rep.int('s', nCols), t)
    v <- c(names(df), v)
    nRows <- nRows + 1L
  }
  
  ## create references
  r <- .Call("openxlsx_ExcelConvertExpand", startCol:(startCol+nCols-1L), LETTERS, as.character(startRow:(startRow+nRows-1L)))
  
  ##Append hyperlinks, convert h to s in cell type
  if("hyperlink" %in% colClasses){
    
    hInds <- which(t == "h")
    
    if(length(hInds) > 0){
      t[hInds] <- "s"
      
      exHlinks <- hyperlinks[[sheet]]
      newHlinks <- r[hInds]
      names(newHlinks) <- replaceIllegalCharacters(v[hInds])
      
      if(!is.null(hlinkNames) & length(hlinkNames) == length(hInds))
        v[hInds] <- hlinkNames 
      
      if(exHlinks[[1]] == ""){
        hyperlinks[[sheet]] <<- newHlinks
      }else{
        allHlinks <- c(exHlinks, newHlinks)
        allHlinks <- allHlinks[!duplicated(allHlinks, fromLast = TRUE)]
        allHlinks <- allHlinks[order(nchar(allHlinks), allHlinks)]
        hyperlinks[[sheet]] <<- allHlinks
      }
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
    v[strFlag] <- match(newStrs, sharedStrings) - 1L
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
  
  dataCount[[sheet]] <<- dataCount[[sheet]] + 1L
  
})          






Workbook$methods(updateStyles2 = function(style){
  
  
  ## Updates styles.xml
  xfNode <- list(numFmtId = 0,
                 fontId = 0,
                 fillId = 0,
                 borderId = 0,
                 xfId = 0)
  
  baseFont <- .self$getBaseFont()
  defaultFontSize <- sprintf('<sz %s=\"%s\"/>', names(baseFont$size), baseFont$size)
  defaultFontColour <- sprintf('<color %s=\"%s\"/>', names(baseFont$colour), baseFont$colour)
  defaultFontName <- sprintf('<name %s=\"%s\"/>', names(baseFont$name), baseFont$name)
  
  style <- style$as.list()
  alignmentFlag <- FALSE
  
  ## Font
  if(!is.null(style$fontName) | !is.null(style$fontSize) |  !is.null(style$fontColour) | !is.null(style$fontDecoration)){
    
    if(!is.null(style$fontSize$val))
      style$fontSize$val <- as.character(style$fontSize$val)
    
    if(!is.null(style$fontDecoration))
      style$fontDecoration <- c("BOLD", "ITALIC", "UNDERLINE", "UNDERLINE2", "STRIKEOUT")[c("BOLD", "ITALIC", "UNDERLINE", "UNDERLINE2", "STRIKEOUT") %in% style$fontDecoration]
    
    fontNode <- .Call("openxlsx_createFontNode", style, defaultFontSize, defaultFontColour, defaultFontName)
    fontId <- which(styles$font == fontNode) - 1L
    
    if(length(fontId) == 0){
      
      fontId <- length(styles$fonts)
      styles$fonts <<- c(styles[["fonts"]], fontNode)        
      
    }
    
    xfNode$fontId <- fontId
    xfNode <- append(xfNode, list("applyFont" = "1"))
  }
  
  
  ## numFmt
  if(!is.null(style$numFmt)){
    if(as.integer(style$numFmt$numFmtId) > 0){
      
      style$numFmt$numFmtId <- as.character(style$numFmt$numFmtId)
      numFmtId <- style$numFmt$numFmtId
      
      if(as.integer(numFmtId) > 163L){
        
        tmp <- style$numFmt$formatCode
        styles$numFmts <<- unique(c(styles$numFmts, sprintf('<numFmt numFmtId="%s" formatCode="%s"/>', numFmtId, tmp)))
        
      }
      
      xfNode$numFmtId <- numFmtId
      xfNode <- append(xfNode, list("applyNumberFormat" = "1"))
      
    }
  }
  
  ## Fill
  if(!is.null(style$fillFg) | !is.null(style$fillBg)){
    
    fillNode <- .Call("openxlsx_createFillNode", style)
    fillId <- which(styles$fill == fillNode) - 1L
    
    if(length(fillId) == 0){      
      fillId <- length(styles$fills)
      styles$fills <<- c(styles$fills, fillNode)
    }
    
    xfNode$fillId <- fillId
    xfNode <- append(xfNode, list("applyFill" = "1"))
  }
  
  ## Border
  if(any(!is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom)))){
    
    borderNode <-  .Call("openxlsx_createBorderNode", style, c("borderLeft", "borderRight", "borderTop", "borderBottom"))
    borderId <- which(styles$borders == borderNode) - 1L
    
    if(length(borderId) == 0){
      borderId <- length(styles$borders)
      styles$borders <<- c(styles$borders, borderNode)
    }
    
    xfNode$borderId <- borderId
    xfNode <- append(xfNode, list("applyBorder" = "1"))
  }
  
  ## Alignment
  if(!is.null(style$halign) | !is.null(style$valign) | !is.null(style$wrapText) | !is.null(style$textRotation)){
    
    if(!is.null(style$textRotation))
      style$textRotation <- as.character(style$textRotation)
    alignmentFlag <- TRUE
    
    alignNode <- .Call("openxlsx_createAlignmentNode", style)
    xfNode <- append(xfNode, list("applyAlignment" = "1"))
  }
  
  if(alignmentFlag){
    xfNode <- paste0("<xf ", paste(paste0(names(xfNode), '="', xfNode, '"'), collapse = " "), ">", alignNode, '</xf>')  
  }else{
    xfNode <- paste0("<xfv", paste(paste0(names(xfNode), '="', xfNode, '"'), collapse = " "), "/>")
  }
  
  styleId <- which(styles$cellXfs == xfNode) - 1L
  if(length(styleId) == 0){
    styleId <- length(styles$cellXfs)
    styles$cellXfs <<- c(styles$cellXfs, xfNode)
  }
  
  
  return(as.integer(styleId))
  
})



Workbook$methods(updateStyles = function(style){
  
  
  ## Updates styles.xml
  xfNode <- list(numFmtId = 0,
                 fontId = 0,
                 fillId = 0,
                 borderId = 0,
                 xfId = 0)
  
  
  alignmentFlag <- FALSE
  
  ## Font
  if(!is.null(style$fontName) |
       !is.null(style$fontSize) |
       !is.null(style$fontColour) |
       !is.null(style$fontDecoration) | 
       !is.null(style$fontFamily) | 
       !is.null(style$fontScheme)){
    
    fontNode <- .self$createFontNode(style)
    fontId <- which(styles$font == fontNode) - 1L
    
    if(length(fontId) == 0){
      
      fontId <- length(styles$fonts)
      styles$fonts <<- append(styles[["fonts"]], fontNode)        
      
    }
    
    xfNode$fontId <- fontId
    xfNode <- append(xfNode, list("applyFont" = "1"))
  }
  
  
  ## numFmt
  if(!is.null(style$numFmt)){
    if(as.integer(style$numFmt$numFmtId) > 0){
      numFmtId <- style$numFmt$numFmtId
      if(as.integer(numFmtId) > 163L){
        
        tmp <- style$numFmt$formatCode
        
        styles$numFmts <<- unique(c(styles$numFmts,
                                    sprintf('<numFmt numFmtId="%s" formatCode="%s"/>', numFmtId, tmp)
        ))
      }
      
      xfNode$numFmtId <- numFmtId
      xfNode <- append(xfNode, list("applyNumberFormat" = "1"))
    }
  }
  
  ## Fill
  if(!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg)){
    
    fillNode <- .self$createFillNode(style)
    fillId <- which(styles$fill == fillNode) - 1L
    
    if(length(fillId) == 0){      
      fillId <- length(styles$fills)
      styles$fills <<- c(styles$fills, fillNode)
    }
    xfNode$fillId <- fillId
    xfNode <- append(xfNode, list("applyFill" = "1"))
  }
  
  ## Border
  if(any(!is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom)))){
    
    borderNode <- .self$createBorderNode(style)
    borderId <- which(styles$borders == borderNode) - 1L
    
    if(length(borderId) == 0){
      borderId <- length(styles$borders)
      styles$borders <<- c(styles$borders, borderNode)
    }
    
    xfNode$borderId <- borderId
    xfNode <- append(xfNode, list("applyBorder" = "1"))
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
    
    if(!is.null(style$wrapText)){
      if(style$wrapText)
        alignNode <- paste(alignNode, 'wrapText="1"')
    }
    
    
    alignNode <- paste0(alignNode, "/>")
    
    alignmentFlag <- TRUE
    xfNode <- append(xfNode, list("applyAlignment" = "1"))
  }
  
  if(alignmentFlag){
    xfNode <- paste0("<xf ", paste(paste0(names(xfNode), '="',xfNode, '"'), collapse = " "), ">", alignNode, '</xf>')  
  }else{
    xfNode <- paste0("<xf ", paste(paste0(names(xfNode), '="',xfNode, '"'), collapse = " "), "/>")
  }
  
  styleId <- which(styles$cellXfs == xfNode) - 1L
  if(length(styleId) == 0){
    styleId <- length(styles$cellXfs)
    styles$cellXfs <<- c(styles$cellXfs, xfNode)
  }
  
  
  return(as.integer(styleId))
  
})


Workbook$methods(getBaseFont = function(){
  
  baseFont <- styles$fonts[[1]]
  
  sz <- getAttrs(baseFont, "<sz ")
  colour <- getAttrs(baseFont, "<color ")
  name <- getAttrs(baseFont, "<name ")  
  
  if(length(sz[[1]]) == 0)
    sz <- list("val" = "10")
  
  if(length(colour[[1]]) == 0)
    colour <- list("rgb" = "#000000")
  
  if(length(name[[1]]) == 0)
    name <- list("val" = "Calibri")
  
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
    fontNode <- paste0(fontNode, '<u val="single"/>')
  
  if("UNDERLINE2" %in% style$fontDecoration)
    fontNode <- paste0(fontNode, '<u val="double"/>')
  
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


Workbook$methods(createFillNode = function(style, patternType="solid"){
  
  fill <- style$fill
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
  
  ## rename styleObjects sheet component
  if(length(styleObjects) > 0){
    
    styleObjects <<- lapply(styleObjects, function(x){
      
      if(x$sheet == oldName)
        x$sheet <- newSheetName
      
      return(x)      
      
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
      ws$sheetViews <- gsub("/></sheetViews>", paste0(">", freezePane[[i]], "</sheetView></sheetViews>"), ws$sheetViews)
    
    if(length(worksheets[[i]]$mergeCells) > 0)
      ws$mergeCells <- paste0(sprintf('<mergeCells count="%s">', length(worksheets[[i]]$mergeCells)), pxml(worksheets[[i]]$mergeCells), '</mergeCells>')
    
    if(length(worksheets[[i]]$conditionalFormatting) > 0){
      nms <- names(ws$conditionalFormatting)
      uNames <- unique(nms)
      ws$conditionalFormatting <- paste(sapply(uNames, function(x) paste0(sprintf('<conditionalFormatting sqref="%s">', x), pxml(ws$conditionalFormatting[nms == x]), '</conditionalFormatting>')), collapse = "")
    }
    
    ## Header footer
    if(!is.null(ws$headerFooter))
      ws$headerFooter <- genHeaderFooterNode(ws$headerFooter)
    
    if(!is.null(ws$sheetPr))
      ws$sheetPr <- ws$sheetPr #paste0("<sheetPr>", paste(ws$sheetPr, collapse = ""), "</sheetPr>")
    
    if(length(worksheets[[i]]$tableParts) > 0)
      ws$tableParts <- paste0(sprintf('<tableParts count="%s">', length(worksheets[[i]]$tableParts)), pxml(worksheets[[i]]$tableParts), '</tableParts>')
    
    if(length(worksheets[[i]]$extLst) > 0)
      ws$extLst <- sprintf('<extLst>%s</extLst>', paste(worksheets[[i]]$extLst, collapse = ""))
    
    if(hyperlinks[[i]][[1]] != ""){
      nTables <- length(tables)
      nHLinks <- length(hyperlinks[[i]])
      hInds <- as.integer(1:nHLinks + 3L + nTables - 1L)
      ws$hyperlinks <- paste0('<hyperlinks>', paste(sprintf('<hyperlink ref="%s" r:id="rId%s"/>', hyperlinks[[i]], hInds), collapse = ""), '</hyperlinks>')  
    }
    
    sheetDataInd <- which(names(ws) == "sheetData")
    prior <- paste0(header, pxml(ws[1:(sheetDataInd - 1L)]))
    post <- paste0(pxml(ws[(sheetDataInd + 1L):length(ws)]), "</worksheet>")
    
    ## Sort sheetData before writing
    if(dataCount[[i]] > 1L | length(rowHeights[[i]]) > 0){
      r <- unlist(lapply(sheetData[[i]], "[[", "r"), use.names = TRUE)     
      sheetData[[i]] <<- sheetData[[i]][order(as.integer(names(r)), nchar(r), r)]
      dataCount[[i]] <<- 1L
        
      if(length(styleInds[[i]]) > 0)
        styleInds[[i]] <<- styleInds[[i]][match(unlist(lapply(sheetData[[i]], "[[", "r"), use.names = FALSE), names(styleInds[[i]]))]
    }
  
    if(length(rowHeights[[i]]) == 0){

      .Call("openxlsx_quickBuildCellXML",
            prior,
            post,
            sheetData[[i]],
            as.integer(names(sheetData[[i]])),
            styleInds[[i]],
            file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)),
            PACKAGE = "openxlsx")
      
    }else{
      
      ## row heights will always be in order and all row heights are given rows in preSaveCleanup  
      .Call("openxlsx_quickBuildCellXML2",
            prior,
            post,
            sheetData[[i]],
            as.integer(names(sheetData[[i]])),
            as.character(unname(styleInds[[i]])),
            unlist(rowHeights[[i]]),
            file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)),
            PACKAGE="openxlsx")
      
    }
    
    ## write worksheet rels
    if(length(worksheets_rels[[i]]) > 0 ){
      
      ws_rels <- worksheets_rels[[i]]
      
      if(hyperlinks[[i]][[1]] != ""){
        
        if("UTF-8" %in% Encoding(hyperlinks[[i]])){
          fromEnc <- "UTF-8"
        }else{
          fromEnc <- ""
        }
        
        ws_rels <- c(ws_rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>', hInds, iconv(names(hyperlinks[[i]]), from = fromEnc, to = "UTF-8")))
        
      }
      
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
  names(widths) <- cols
  
  autoColsInds <- widths == "auto"
  autoCols <- cols[autoColsInds]
  
  ## If any not auto
  if(any(!autoColsInds))
    widths[!autoColsInds] <- as.numeric(widths[!autoColsInds]) + 0.71
  
  ## If any auto
  if(length(autoCols) > 0){
    
    ## only run if data on worksheet
    if(length(sheetData[[sheet]]) == 0){
      
      missingAuto <- autoCols
      
    }else{
      
      ## First thing - get base font max character width
      baseFont <- getBaseFont()
      baseFontName <- unlist(baseFont$name, use.names = FALSE)
      if(is.null(baseFontName)){
        baseFontName <- "calibri"
      }else{
        baseFontName <- gsub(" ", ".", tolower(baseFontName))
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
      allCharWidths <- rep(baseFontCharWidth, length(sheetData[[sheet]]))
      #########----------------------------------------------------------------

      ## get char widths for each style object
      if(length(styleObjects) > 0){
        
        thisSheetName <- names(worksheets)[sheet]
        
        ## Calc font width for all styles on this worksheet
        styleIds <- styleInds[[sheet]]
        styObSubet <- styleObjects[sort(unique(styleIds))]
        stySubset <- lapply(styObSubet, "[[", "style") 
        
        ## loop through stlye objects assignin a charWidth else baseFontCharWidth
        styleCharWidths <- unlist(lapply(stySubset, function(thisStyle){
          
          fN <- unlist(thisStyle$fontName, use.names = FALSE)
          if(is.null(fN)){
            fN <- "calibri"
          }else{
            fN <- gsub(" ", ".", tolower(fN))
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
          
          styleMaxCharWidth
          
        }), use.names = FALSE)
        
        
        ## Now assign all cells a character width
        allCharWidths <- styleCharWidths[styleInds[[sheet]]]
        allCharWidths[is.na(allCharWidths)] <- baseFontCharWidth
      }
      
      ## Now that we have the max character width for the largest font on the page calculate the column widths
      calculatedWidths <- .Call("openxlsx_calcColumnWidths",
                                sheetData = sheetData[[sheet]],
                                sharedStrings = unlist(sharedStrings, use.names = FALSE),
                                columnInds = as.integer(autoCols),
                                width = allCharWidths,
                                baseFontCharWidth = baseFontCharWidth,
                                minW = getOption("openxlsx.minWidth", 3),
                                maxW = getOption("openxlsx.maxWidth", 250))
      
      missingAuto <- autoCols[!autoCols %in% names(calculatedWidths)]
      widths[names(calculatedWidths)] <- calculatedWidths + 0.71
      
    }
    
    widths[missingAuto] <- 9.15
    
  }
  
  ## Calculate width of auto
  colNodes <- sprintf('<col min="%s" max="%s" width="%s" customWidth="1"/>', cols, cols, widths)
  
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
  allRowHeights <- allRowHeights[order(as.integer(names(allRowHeights)))]
  
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
  nSheets <- length(unlist(sheetNames, use.names = FALSE))
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
  sheetOrder <<- c(sheetOrder[sheetOrder < sheet], sheetOrder[sheetOrder > sheet] - 1L)
  
  ## remove styleObjects
  if(length(styleObjects) > 0)
    styleObjects <<-styleObjects[unlist(lapply(styleObjects, "[[", "sheet"), use.names = FALSE) != sheetName]
  
  ## wont't remove tables and then won't need to reassign table r:id's
  worksheets[[sheet]] <<- NULL
  worksheets_rels[[sheet]] <<- NULL
  
  ## drawing will always be the first relationship and printerSettings second
  for(i in 1:(nSheets-1L))
    worksheets_rels[[i]][1:2] <<- genBaseSheetRels(i)
  
  
  workbook$sheets <<- workbook$sheets[!grepl(sprintf('name="%s"', sheetName), workbook$sheets)]
  ## Reset rIds
  for(i in (sheet+1L):nSheets)
    workbook$sheets <<- gsub(paste0("rId", i), paste0("rId", i-1L), workbook$sheets)
  
  ## Can remove highest sheet
  workbook.xml.rels <<- workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), workbook.xml.rels)]
  freezePane[[sheet]] <<- NULL
  dataCount[[sheet]] <<- NULL
  
  invisible(1)
  
})


Workbook$methods(addDXFS = function(style){
  
  dxf <- '<dxf>'
  dxf <- paste0(dxf, createFontNode(style))
  fillNode <- NULL
  
  if(!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg))
    dxf <- paste0(dxf, createFillNode(style))
  
  if(any(!is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom))))
    dxf <- paste0(dxf, createBorderNode(style))
  
  dxf <- paste(dxf, "</dxf>")
  if(dxf %in% styles$dxfs)
    return(which(styles$dxfs == dxf) - 1L)
  
  dxfId <- length(styles$dxfs)
  styles$dxfs <<- c(styles$dxfs, dxf)
  
  return(dxfId)
})



Workbook$methods(conditionalFormatCell = function(sheet, startRow, endRow, startCol, endCol, dxfId, formula, type){
  
  sheet = validateSheet(sheet)
  sqref <- paste(getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol))), collapse = ":")
  
  ## Increment priority of conditional formatting rule
  if(length((worksheets[[sheet]]$conditionalFormatting)) > 0){
    for(i in length(worksheets[[sheet]]$conditionalFormatting):1)
      worksheets[[sheet]]$conditionalFormatting[[i]] <<- gsub('(?<=priority=")[0-9]+', i+1L, worksheets[[sheet]]$conditionalFormatting[[i]], perl = TRUE)
  }
  
  nms <- c(names(worksheets[[sheet]]$conditionalFormatting), sqref)
  
  if(type == "expression"){
    
    cfRule <- sprintf('<cfRule type="expression" dxfId="%s" priority="1"><formula>%s</formula></cfRule>', dxfId, formula)
    
  }else if(type == "dataBar"){
    
    if(length(formula) == 2){
      negColour <- formula[[1]]
      posColour <- formula[[2]]
    }else{
      posColour <- formula
      negColour <- "FFFF0000"
    }
    
    guid <- paste0("F7189283-14F7-4DE0-9601-54DE9DB", 40000L + length(worksheets[[sheet]]$extLst))
    cfRule <- sprintf('<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="%s"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{%s}</x14:id></ext></extLst></cfRule>', posColour, guid)
    worksheets[[sheet]]$extLst <<- c(worksheets[[sheet]]$extLst, genExtLst(guid, sqref, posColour, negColour))
    
  }else if(length(formula) == 2L){
    
    cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="%s"/><color rgb="%s"/></colorScale></cfRule>', formula[[1]], formula[[2]])
    
  }else{
    
    cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="%s"/><color rgb="%s"/><color rgb="%s"/></colorScale></cfRule>', formula[[1]], formula[[2]], formula[[3]])
  }
  
  worksheets[[sheet]]$conditionalFormatting <<- append(worksheets[[sheet]]$conditionalFormatting, cfRule)              
  
  names(worksheets[[sheet]]$conditionalFormatting) <<- nms
  
  invisible(0)
  
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
  
  ## Remove intersection
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
      attrs <- c(attrs, sprintf('ySplit="%s"', firstActiveRow - 1L))
    
    if(firstActiveCol > 1)
      attrs <- c(attrs, sprintf('xSplit="%s"', firstActiveCol - 1L))
    
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
  
  imageNo <- length((drawings[[sheet]])) + 1L
  mediaNo <- length(media) + 1L
  
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
  </xdr:from>' , startCol - 1L, colOffset, startRow - 1L, rowOffset)
  
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
  nPivots <- length(pivotDefinitions)
  
  ## add a worksheet if none added
  if(nSheets == 0){
    warning("Workbook does not contain any worksheets. A worksheet will be added.")
    .self$addWorksheet("Sheet 1")
    nSheets <- 1L
  }
  
  ## Assign relationship Ids  
  nChildren <- 1:length(workbook.xml.rels)
  splitRels <- strsplit(workbook.xml.rels, split = " ")
  nms <- lapply(splitRels, function(x) gsub("(.*)=(.*)", "\\1", x[2:length(x)]))
  rAttrs <- lapply(splitRels, function(x) gsub("/>$", "", gsub('(.*)="(.*)"', "\\2", x[2:length(x)])))   
  rAttrs <- lapply(nChildren, function(i) {names(rAttrs[[i]]) <- nms[[i]]; return(rAttrs[[i]])})
  
  rIds <- paste0("rId", nChildren)
  targets <- unlist(lapply(rAttrs, "[[", "Target"))
  
  ## Reorderinf of workbook.xml.rels
  sheetInds <- which(targets %in% sprintf("worksheets/sheet%s.xml", 1:nSheets))
  sheetNumbers <- as.integer(gsub("[^0-9]", "", targets[sheetInds], perl = TRUE))
  sheetInds <- sheetInds[order(sheetNumbers)]  ## make sure sheets will be in order
  
  ## get index of each child element for ordering
  stylesInd <- which(targets == "styles.xml")
  themeInd <- which(targets %in% sprintf("theme/theme%s.xml", 1:nThemes))
  connectionsInd <- which(targets %in% "connections.xml")
  
  extRefInds <- which(targets %in% sprintf("externalLinks/externalLink%s.xml", 1:nExtRefs))
  sharedStringsInd <- which(targets == "sharedStrings.xml")
  tableInds <- which(grepl("table[0-9]+.xml", targets))
  pivotNode <- workbook.xml.rels[which(targets %in% sprintf("pivotCache/pivotCacheDefinition%s.xml", 1:nPivots))] ## don't want to re-assign rIds for pivot tables
  
  ## Reorder children of workbook.xml.rels
  workbook.xml.rels <<- workbook.xml.rels[c(sheetInds, extRefInds, themeInd, connectionsInd, stylesInd, sharedStringsInd, tableInds)]
  
  ## Re assign rIds to children of workbook.xml.rels
  workbook.xml.rels <<- unlist(lapply(1:length(workbook.xml.rels), function(i) {
    gsub('(?<=Relationship Id="rId)[0-9]+', i, workbook.xml.rels[[i]], perl = TRUE)
  }))
  
  workbook.xml.rels <<- c(workbook.xml.rels, pivotNode)
  
  if(!is.null(vbaProject))
    workbook.xml.rels <<- c(workbook.xml.rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>', 1L + length(workbook.xml.rels)))
  
  
  ## Reassign rId to workbook sheet elements, (order sheets by sheetId first)
  sId <- as.integer(unlist(regmatches(workbook$sheets, gregexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE))))
  workbook$sheets <<- sapply(order(sId), function(i) gsub('(?<=id="rId)[0-9]+', i, workbook$sheets[[i]], perl = TRUE))
  workbook$sheets <<- sapply(1:nSheets, function(i) gsub('(?<=sheetId=")[0-9]+', i, workbook$sheets[[i]], perl = TRUE))
  if(!is.null(sheetOrder))
    workbook$sheets <<- workbook$sheets[sheetOrder]
  
  ## update workbook r:id to match reordered workbook.xml.rels externalLink element
  if(length(extRefInds) > 0){
    newInds <- as.integer(1:length(extRefInds) + length(sheetInds))
    workbook$externalReferences <<- paste0("<externalReferences>",
                                           paste0(sprintf('<externalReference r:id=\"rId%s\"/>', newInds), collapse = ""),
                                           "</externalReferences>")
  }
  
  ## styles
  numFmtIds <- 50000L
  styleInds <<- lapply(1:length(worksheets), function(i){
    y <- rep.int(NA, length(sheetData[[i]]))
    names(y) <- unlist(lapply(sheetData[[i]], "[[", "r"), use.names = FALSE)
    return(y)
  })
  
  for(x in styleObjects){
    if(length(x$rows) > 0 & length(x$cols) > 0){
      
      this.sty <- x$style$copy()  
      
      if(!is.null(this.sty$numFmt)){
        if(this.sty$numFmt$numFmtId == 9999){
          this.sty$numFmt$numFmtId <- numFmtIds
          numFmtIds <- numFmtIds + 1L
        }
      }
      
      ## convert sheet name to index
      sheet <- which(names(worksheets) == x$sheet)
      sId <- .self$updateStyles(this.sty) ## this creates the XML for styles.XML
      
      ## In here we create any styleInds that don't yet have a sheetData
      refsToStyle <- paste0(.Call('openxlsx_convert2ExcelRef', PACKAGE = 'openxlsx', x$cols, LETTERS), x$rows)
      styleInds[[sheet]][names(styleInds[[sheet]]) %in% refsToStyle] <<- sId
      
      toAppend <- refsToStyle[!refsToStyle %in% names(styleInds[[sheet]])]
      if(length(toAppend) > 0){
        tmp <- rep.int(sId, length(toAppend))
        names(tmp) <- toAppend
        styleInds[[sheet]] <<- c(styleInds[[sheet]], tmp)
        
        ## Now append new cells to sheetData - incremenet dataCount so sorting will take place
        newCells <- .Call("openxlsx_buildCellList", toAppend , rep(as.character(NA), length(toAppend)) , rep(as.character(NA), length(toAppend)), PACKAGE="openxlsx")
        names(newCells) <- gsub("[A-Z]", "", toAppend)
        sheetData[[sheet]] <<- append(sheetData[[sheet]], newCells)
        dataCount[[sheet]] <<- dataCount[[sheet]] + 1L
        
      }
      
    }
        
  }
  
  
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



Workbook$methods(addStyle = function(sheet, style, rows, cols, stack){
  
  sheet <- names(worksheets)[[sheet]]
  
  if(length(styleObjects) == 0){
    
    styleObjects <<- list(list(style = style,
                               sheet =  sheet,
                               rows = rows,
                               cols = cols))   
  }else if(stack){
    
    
    nStyles <- length(styleObjects)
    
    ## ********** Assume all styleObjects cells have one a single worksheet **********
    ## Loop through existing styleObjects
    newInds <- 1:length(rows)
    for(i in 1:nStyles){
      
      if(sheet == styleObjects[[i]]$sheet){
        
        ## Now check rows and cols intersect
        
        toRemoveInds <- which(match(styleObjects[[i]]$rows, rows) == match(styleObjects[[i]]$cols, cols))
        mergeInds <- which(rows %in% styleObjects[[i]]$rows & cols %in% styleObjects[[i]]$cols)
        newInds <- newInds[!newInds %in% mergeInds]
        
        if(length(toRemoveInds) > 0){
          
          ## remove these from style object
          styleObjects[[i]]$rows <<- styleObjects[[i]]$rows[-toRemoveInds]
          styleObjects[[i]]$cols <<- styleObjects[[i]]$cols[-toRemoveInds]
          
        }
        
        ## append style object for intersecting cells
        if(length(mergeInds) > 0){
          
          styleObjects <<- append(styleObjects, list(list(style = mergeStyle(styleObjects[[i]]$style, newStyle = style),
                                                          sheet =  sheet,
                                                          rows = rows[mergeInds],
                                                          cols = cols[mergeInds])))        
        }
        
        
      } ## if sheet == styleObjects[[i]]$sheet
      
      
    } ## End of loop through styles
    
    ## append style object for non-intersecting cells
    if(length(newInds) > 0){
      
      styleObjects <<- append(styleObjects, list(list(style = style,
                                                      sheet =  sheet,
                                                      rows = rows[newInds],
                                                      cols = cols[newInds])))
      
    }
    
    
  }else{
    
    styleObjects <<- append(styleObjects, list(list(style = style,
                                                    sheet =  sheet,
                                                    rows = rows,
                                                    cols = cols)))
    
    
  } ## End if(length(styleObjects) > 0) else if(stack) {}
  
  
  
  
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
    showText <- c(showText, "\nWorksheets:\n", "No worksheets attached\n")
  }
  
  ## images
  if(nImages > 0)
    showText <- c(showText, "\nImages:\n", sprintf('Image %s: "%s"\n', 1:nImages, media))
  
  if(nCharts > 0)
    showText <- c(showText, "\nCharts:\n", sprintf('Chart %s: "%s"\n', 1:nImages, media))
  
  if(nSheets > 0)
    showText <- c(showText, sprintf("Worksheet write order: %s", paste(sheetOrder, collapse = ", ")))
  
  cat(unlist(showText))
  
})


# ## function to create the below
# strs <- "data.frame("
# for(i in 1:nrow(tab))
#   strs <- append(strs, paste0('"', gsub(" ", ".", tolower(tab$Font[[i]])), '" = ',  paste(capture.output(dput(unname(unlist(tab[1,2:ncol(tab)])))), collapse = ""), ", \n"))
# strs[length(strs)] <- gsub(", \\\n$", ")", strs[length(strs)])
# cat(strs)



## Character width lookup table
openxlsxFontSizeLookupTable <- 
  data.frame( "agency.fb"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.25, 2.42875, 2.7325, 2.7325, 2.875, 2.875, 3.07125, 3.07125),
              "aharoni"= c(0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 2.875),
              "algerian"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "andalus"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "angsana.new"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25),
              "angsanaupc"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25),
              "aparajita"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "arial.black"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625),
              "arial.narrow"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375, 3.21375),
              "arial.rounded.mt.bold"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "arial.unicode.ms"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "arial"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "calibri.light"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "calibri"= c(0.71375, 0.8575, 0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "californian.fb"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "calisto.mt"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "cambria.math"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "cambria"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "candara"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "castellar"= c(1.17875, 1.17875, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125, 4.6075, 4.7675, 4.7675, 5.08875, 5.25, 5.25),
              "centaur"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "century.gothic"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "century.schoolbook"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "century"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "chiller"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "colonna.mt"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "comic.sans.ms"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "consolas"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "constantia"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "cooper.black"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "copperplate.gothic.bold"= c(1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.32125, 4.32125, 4.6075, 4.6075, 4.7675, 5.08875, 5.08875, 5.25),
              "copperplate.gothic.light"= c(1.17875, 1.17875, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.7675, 5.08875, 5.25, 5.25),
              "corbel"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "cordia.new"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "david"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375),
              "dfkai-sb"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "dilleniaupc"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075),
              "dokchampa"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "dotum"= c(0.71375, 0.71375, 1.17875, 1, 1, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125),
              "dotumche"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ebrima"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "edwardian.script.itc"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.21375, 3.3575, 3.3575),
              "elephant"= c(1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.5, 3.5, 3.69625, 3.9825, 3.9825, 4.1425, 4.46375, 4.46375, 4.6075, 4.94625, 4.94625, 5.08875, 5.3925, 5.3925, 5.57125),
              "engravers.mt"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675),
              "eras.bold.itc"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125, 4.6075, 4.6075, 4.7675),
              "eras.demi.itc"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "eras.light.itc"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "eras.medium.itc"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425),
              "estrangelo.edessa"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "eucrosiaupc"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "euphemia"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125),
              "fangsong"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "felix.titling"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125),
              "fixedsys"= c(1, 1, 1, 1, 1, 1, 1, 1, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 3.5, 3.5, 3.5, 3.5, 3.5, 3.5, 3.5, 3.5, 3.5, 4.7675, 4.7675, 4.7675, 4.7675),
              "footlight.mt.light"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "forte"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "franklin.gothic.book"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "franklin.gothic.demi.cond"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "franklin.gothic.demi"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "franklin.gothic.heavy"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "franklin.gothic.medium.cond"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625),
              "franklin.gothic.medium"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "frankruehl"= c(0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 2.875),
              "freesiaupc"= c(0.57125, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375),
              "freestyle.script"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875),
              "french.script.mt"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "gabriola"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "gadugi"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "gautami"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "georgia"= c(1, 1.17875, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125, 4.46375),
              "gigi"= c(0.8575, 0.8575, 0.8575, 0.8575, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.625, 1.625, 1.96375, 1.96375, 1.96375, 2.42875, 2.42875, 2.42875, 2.7325, 2.7325, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.5, 3.83875),
              "gill.sans.mt.condensed"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "gill.sans.mt.ext.condensed.bold"= c(0.3925, 0.3925, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875),
              "gill.sans.mt"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "gill.sans.ultra.bold.condensed"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.1425),
              "gill.sans.ultra.bold"= c(1.32125, 1.4825, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.875, 3.07125, 3.07125, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.57125, 5.71375, 5.71375, 6.0175, 6.19625, 6.3575),
              "gisha"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "gloucester.mt.extra.condensed"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325),
              "goudy.old.style"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "goudy.stout"= c(1.32125, 1.4825, 1.625, 1.80375, 1.96375, 2.25, 2.25, 2.42875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.83875, 3.83875, 3.9825, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925, 5.57125, 5.875, 6.0175, 6.0175),
              "gulim"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.80375, 1.80375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "gulimche"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "gungsuh"= c(0.71375, 0.71375, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125),
              "gungsuhche"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "haettenschweiler"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375),
              "harlow.solid.italic"= c(0.785, 0.785, 0.92875, 0.92875, 1.1075, 1.285, 1.285, 1.42875, 1.625, 1.625, 1.7675, 1.94625, 1.94625, 2.1425, 2.1425, 2.285, 2.285, 2.46375, 2.66, 2.66, 2.80375, 2.9825, 2.9825, 3.125, 3.32125, 3.32125, 3.5, 3.5, 3.6425),
              "harrington"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "high.tower.text"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5),
              "impact"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825),
              "imprint.mt.shadow"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "informal.roman"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.69625, 3.83875),
              "irisupc"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325),
              "iskoola.pota"= c(0.71375, 0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "jasmineupc"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375, 3.21375),
              "jokerman"= c(1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.32125, 4.32125, 4.6075, 4.6075, 4.7675, 5.08875, 5.08875, 5.25),
              "juice.itc"= c(0.3925, 0.57125, 0.71375, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.58875),
              "kaiti"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "kalinga"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825),
              "kartika"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "khmer.ui"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "kodchiangupc"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.42875),
              "kristen.itc"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "kunstler.script"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.42875),
              "lao.ui"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "latha"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "leelawadee"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "levenim.mt"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "lilyupc"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375),
              "lucida.bright"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "lucida.calligraphy"= c(1.07125, 1.25, 1.42875, 1.58875, 1.58875, 1.91, 1.94625, 2.1075, 2.25, 2.3925, 2.58875, 2.7675, 2.91, 2.91, 3.285, 3.285, 3.42875, 3.57125, 3.7675, 3.94625, 4.08875, 4.2325, 4.285, 4.6075, 4.6075, 4.75, 4.94625, 5.08875, 5.2675),
              "lucida.console"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "lucida.fax"= c(1, 1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075),
              "lucida.handwriting"= c(1.07125, 1.25, 1.42875, 1.58875, 1.58875, 1.91, 2.1075, 2.1075, 2.3925, 2.3925, 2.625, 2.7675, 2.91, 3.05375, 3.285, 3.42875, 3.42875, 3.7325, 3.7675, 3.94625, 4.08875, 4.2325, 4.42875, 4.6075, 4.75, 4.75, 5.08875, 5.08875, 5.2675),
              "lucida.sans.typewriter"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "lucida.sans.unicode"= c(1, 1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075),
              "lucida.sans"= c(1, 1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075),
              "magneto"= c(1.17875, 1.32125, 1.625, 1.625, 1.80375, 2.1075, 2.1075, 2.25, 2.58875, 2.7325, 2.7325, 3.07125, 3.21375, 3.21375, 3.5, 3.69625, 3.69625, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925, 5.57125, 5.71375),
              "maiandra.gd"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "malgun.gothic"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "mangal"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "marlett"= c(1.625, 1.80375, 2.1075, 2.25, 2.42875, 2.7325, 2.875, 3.07125, 3.3575, 3.5, 3.69625, 3.9825, 4.1425, 4.32125, 4.6075, 4.7675, 4.94625, 5.25, 5.3925, 5.57125, 5.875, 6.0175, 6.19625, 6.5, 6.6425, 6.82125, 7.125, 7.2675, 7.46375),
              "matura.mt.script.capitals"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.7675),
              "meiryo.ui"= c(0.8575, 1, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "meiryo"= c(0.8575, 1, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "microsoft.himalaya"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25),
              "microsoft.jhenghei.ui"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.1425),
              "microsoft.jhenghei"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.1425),
              "microsoft.new.tai.lue"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "microsoft.phagspa"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "microsoft.sans.serif"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "microsoft.tai.le"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "microsoft.uighur"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "microsoft.yahei.ui"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "microsoft.yahei"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "microsoft.yi.baiti"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "miriam.fixed"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "miriam"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "mistral"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375),
              "modern.no..20"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 2.875, 3.07125, 3.21375, 3.21375),
              "modern"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "mongolian.baiti"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "monotype.corsiva"= c(0.6425, 0.785, 0.92875, 0.92875, 0.96375, 1.1075, 1.285, 1.285, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.9825, 2.1425, 2.17875, 2.32125, 2.32125, 2.46375, 2.46375, 2.69625, 2.83875, 2.83875, 2.9825, 3.0175, 3.17875, 3.3575, 3.3575, 3.535),
              "moolboran"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875),
              "ms.gothic"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ms.mincho"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ms.outlook"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ms.pgothic"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ms.pmincho"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ms.reference.sans.serif"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.6075),
              "ms.reference.specialty"= c(1.96375, 2.25, 2.58875, 2.7325, 2.875, 3.3575, 3.5, 3.69625, 4.1425, 4.32125, 4.46375, 4.7675, 5.08875, 5.25, 5.57125, 5.71375, 6.0175, 6.3575, 6.5, 6.6425, 7.125, 7.2675, 7.46375, 7.8925, 8.08875, 8.2325, 8.535, 8.8575, 9),
              "ms.sans.serif"= c(0.71375, 0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 1.96375, 2.58875, 2.58875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.07125, 3.07125, 3.07125, 3.5, 3.5, 3.5, 3.83875, 3.83875),
              "ms.serif"= c(0.57125, 0.71375, 0.71375, 1, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.625, 1.625, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.25, 2.58875, 2.58875, 2.58875, 3.69625, 3.69625, 3.69625, 2.875, 3.69625, 3.69625, 3.5, 3.5),
              "ms.ui.gothic"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "mt.extra"= c(1.625, 1.80375, 2.1075, 2.25, 2.42875, 2.7325, 2.875, 3.07125, 3.3575, 3.5, 3.69625, 4.1425, 4.32125, 4.46375, 4.7675, 4.94625, 5.08875, 5.3925, 5.57125, 5.71375, 6.0175, 6.19625, 6.3575, 6.6425, 6.82125, 6.9825, 7.2675, 7.46375, 7.6075),
              "mv.boli"= c(1, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.08875),
              "narkisim"= c(0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 2.875),
              "niagara.engraved"= c(0.3925, 0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375),
              "niagara.solid"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375),
              "nirmala.ui"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "nsimsun"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "nyala"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "ocr.a.extended"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "old.english.text.mt"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "onyx"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375, 1.96375),
              "palace.script.mt"= c(0.42875, 0.42875, 0.6425, 0.6425, 0.6425, 0.785, 0.785, 0.92875, 0.96375, 1.1075, 1.1075, 1.1075, 1.285, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.66, 1.80375, 1.80375, 1.9825, 1.9825, 1.9825, 2.17875, 2.17875, 2.32125, 2.32125, 2.32125),
              "palatino.linotype"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "papyrus"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "parchment"= c(0.25, 0.25, 0.25, 0.25, 0.3925, 0.3925, 0.3925, 0.3925, 0.57125, 0.57125, 0.57125, 0.71375, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 0.8575, 1, 1, 1, 1, 1.17875, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125),
              "perpetua.titling.mt"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625),
              "perpetua"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.21375),
              "plantagenet.cherokee"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "playbill"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.42875),
              "pmingliu-extb"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "pmingliu"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "poor.richard"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "pristina"= c(0.57125, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5),
              "raavi"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.7675),
              "rage.italic"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "rockwell.condensed"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "rockwell.extra.bold"= c(1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875),
              "rockwell"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825),
              "rod"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "roman"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "sakkal.majalla"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "script.mt.bold"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "script"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "segoe.print"= c(1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 2.1075, 2.1075, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.3575, 3.3575, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925),
              "segoe.script"= c(1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 2.1075, 2.1075, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.3575, 3.3575, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925),
              "segoe.ui.light"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "segoe.ui.semibold"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "segoe.ui.semilight"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "segoe.ui.symbol"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "segoe.ui"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "shonar.bangla"= c(0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.42875, 2.7325, 2.7325, 2.7325, 2.875, 2.875, 3.07125),
              "showcard.gothic"= c(1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.32125),
              "shruti"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "simhei"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "simplified.arabic.fixed"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "simplified.arabic"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "simsun-extb"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "simsun"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "small.fonts"= c(0.71375, 0.71375, 1, 1.32125, 1.32125, 1.32125, 1.625, 1.625, 1.625, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.58875, 2.875, 2.58875, 2.58875, 2.875, 2.875, 2.875, 3.5, 3.5, 3.5, 3.5, 3.69625, 3.69625, 3.69625, 3.69625),
              "snap.itc"= c(1.32125, 1.4825, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 5.08875, 5.25, 5.3925, 5.57125, 5.71375, 5.875, 6.19625, 6.3575, 6.3575),
              "sylfaen"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "symbol"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "tahoma"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "tempus.sans.itc"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625),
              "terminal"= c(1, 1, 1, 1.625, 1.625, 1.32125, 1.32125, 1.32125, 1.32125, 2.25, 2.25, 2.25, 2.25, 2.25, 3.5, 3.5, 3.5, 3.5, 2.875, 2.875, 2.875, 4.46375, 4.46375, 4.46375, 4.46375, 4.46375, 5.3925, 5.3925, 5.3925),
              "times.new.roman"= c(0.71375, 0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "traditional.arabic"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "trebuchet.ms"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "tunga"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "tw.cen.mt.condensed.extra.bold"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "tw.cen.mt.condensed"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "tw.cen.mt"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "utsaah"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "vani"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125, 4.46375),
              "verdana"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.6075),
              "vijaya"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "viner.hand.itc"= c(1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 4.94625),
              "vivaldi"= c(0.6425, 0.785, 0.785, 0.92875, 0.96375, 1.1075, 1.285, 1.285, 1.4825, 1.4825, 1.625, 1.7675, 1.80375, 1.80375, 1.9825, 2.1425, 2.17875, 2.32125, 2.32125, 2.46375, 2.69625, 2.69625, 2.69625, 2.83875, 3.0175, 3.0175, 3.17875, 3.17875, 3.3575),
              "vladimir.script"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575),
              "vrinda"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "webdings"= c(1.625, 1.80375, 2.1075, 2.25, 2.42875, 2.7325, 2.875, 3.07125, 3.3575, 3.5, 3.69625, 3.9825, 4.1425, 4.32125, 4.6075, 4.7675, 4.94625, 5.25, 5.3925, 5.57125, 5.875, 6.0175, 6.19625, 6.5, 6.6425, 6.82125, 7.125, 7.2675, 7.46375),
              "wide.latin"= c(2.1075, 2.25, 2.7325, 2.875, 3.07125, 3.5, 3.69625, 3.83875, 4.1425, 4.46375, 4.6075, 4.94625, 5.25, 5.3925, 5.71375, 6.0175, 6.19625, 6.5, 6.82125, 6.9825, 7.2675, 7.46375, 7.75, 8.08875, 8.2325, 8.535, 8.8575, 9, 9.33875)) 


openxlsxFontSizeLookupTableBold <- 
  data.frame( "agency.fb"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.21375, 3.5),
              "aharoni"= c(0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 2.875),
              "algerian"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "andalus"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "angsana.new"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25),
              "angsanaupc"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25),
              "aparajita"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "arial"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "arial.black"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625),
              "arial.narrow"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375, 3.21375),
              "arial.rounded.mt.bold"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "arial.unicode.ms"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "calibri"= c(0.71375, 0.8575, 0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "calibri.light"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "californian.fb"= c(0.8575, 1, 1, 1.17875, 1.17875, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.21375, 3.21375, 3.5, 3.69625, 3.69625, 3.69625, 3.83875, 4.1425, 4.1425),
              "calisto.mt"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "cambria"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.32125),
              "cambria.math"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "candara"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "castellar"= c(1.32125, 1.32125, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375, 4.7675, 4.94625, 4.94625, 5.25, 5.3925, 5.3925),
              "centaur"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "century"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "century.gothic"= c(0.8575, 0.8575, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825),
              "century.schoolbook"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "chiller"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "colonna.mt"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "comic.sans.ms"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "consolas"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "constantia"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "cooper.black"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "copperplate.gothic.bold"= c(1.32125, 1.32125, 1.625, 1.625, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.46375, 4.46375, 4.7675, 4.7675, 4.94625, 5.25, 5.25, 5.3925),
              "copperplate.gothic.light"= c(1.32125, 1.32125, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 4.94625, 5.25, 5.3925, 5.3925),
              "corbel"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "cordia.new"= c(0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "david"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375),
              "dfkai-sb"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "dilleniaupc"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075),
              "dokchampa"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "dotum"= c(0.8575, 0.8575, 1.32125, 1.17875, 1.17875, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "dotumche"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ebrima"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "edwardian.script.itc"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.3575, 3.5, 3.5),
              "elephant"= c(1.32125, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.69625, 3.69625, 3.83875, 4.1425, 4.1425, 4.32125, 4.6075, 4.6075, 4.7675, 5.08875, 5.08875, 5.25, 5.57125, 5.57125, 5.71375),
              "engravers.mt"= c(1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625),
              "eras.bold.itc"= c(1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375, 4.7675, 4.7675, 4.94625),
              "eras.demi.itc"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075),
              "eras.light.itc"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825),
              "eras.medium.itc"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.1425, 4.32125),
              "estrangelo.edessa"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "eucrosiaupc"= c(0.57125, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 2.875, 3.07125, 3.21375, 3.21375),
              "euphemia"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "fangsong"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "felix.titling"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "fixedsys"= c(1.21375, 1.21375, 1.21375, 1.21375, 1.21375, 1.21375, 1.21375, 1.21375, 2.46375, 2.46375, 2.46375, 2.46375, 2.46375, 2.46375, 2.46375, 2.46375, 3.7325, 3.7325, 3.7325, 3.7325, 3.7325, 3.7325, 3.7325, 3.7325, 3.7325, 4.9825, 4.9825, 4.9825, 4.9825),
              "footlight.mt.light"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "forte"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "franklin.gothic.book"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "franklin.gothic.demi"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "franklin.gothic.demi.cond"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "franklin.gothic.heavy"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "franklin.gothic.medium"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "franklin.gothic.medium.cond"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875),
              "frankruehl"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.07125),
              "freesiaupc"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.21375),
              "freestyle.script"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325),
              "french.script.mt"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325),
              "gabriola"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.07125),
              "gadugi"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "gautami"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "georgia"= c(1.17875, 1.32125, 1.32125, 1.32125, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.08875),
              "gigi"= c(1, 1, 1, 1, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.80375, 1.80375, 2.1075, 2.1075, 2.1075, 2.58875, 2.58875, 2.58875, 2.875, 2.875, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.69625, 3.9825),
              "gill.sans.mt"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "gill.sans.mt.condensed"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325),
              "gill.sans.mt.ext.condensed.bold"= c(0.57125, 0.57125, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875),
              "gill.sans.ultra.bold"= c(1.4825, 1.625, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 3.07125, 3.21375, 3.21375, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925, 5.71375, 5.875, 5.875, 6.19625, 6.3575, 6.5),
              "gill.sans.ultra.bold.condensed"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.32125),
              "gisha"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "gloucester.mt.extra.condensed"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.58875, 2.7325, 2.875, 2.875),
              "goudy.old.style"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "goudy.stout"= c(1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.42875, 2.42875, 2.58875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.9825, 3.9825, 4.1425, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925, 5.57125, 5.71375, 6.0175, 6.19625, 6.19625),
              "gulim"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.625, 1.96375, 1.96375, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "gulimche"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "gungsuh"= c(0.8575, 0.8575, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "gungsuhche"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "haettenschweiler"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575),
              "harlow.solid.italic"= c(0.92875, 0.92875, 1.07125, 1.07125, 1.285, 1.42875, 1.42875, 1.58875, 1.7675, 1.7675, 1.94625, 2.1075, 2.1075, 2.285, 2.285, 2.42875, 2.42875, 2.66, 2.80375, 2.80375, 2.94625, 3.125, 3.125, 3.32125, 3.46375, 3.46375, 3.6425, 3.6425, 3.80375),
              "harrington"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "high.tower.text"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "impact"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425),
              "imprint.mt.shadow"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "informal.roman"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.83875, 3.9825),
              "irisupc"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575),
              "iskoola.pota"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875),
              "jasmineupc"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5),
              "jokerman"= c(1.32125, 1.32125, 1.625, 1.625, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.46375, 4.46375, 4.7675, 4.7675, 4.94625, 5.25, 5.25, 5.3925),
              "juice.itc"= c(0.57125, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.7325),
              "kaiti"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "kalinga"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125),
              "kartika"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "khmer.ui"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "kodchiangupc"= c(0.3925, 0.3925, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.42875),
              "kristen.itc"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "kunstler.script"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.58875),
              "lao.ui"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "latha"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "leelawadee"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "levenim.mt"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.32125),
              "lilyupc"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5),
              "lucida.bright"= c(0.8575, 1, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075),
              "lucida.calligraphy"= c(1.25, 1.3925, 1.58875, 1.7325, 1.7325, 2.07125, 2.1075, 2.25, 2.3925, 2.58875, 2.7325, 2.91, 3.05375, 3.05375, 3.42875, 3.42875, 3.57125, 3.7325, 3.94625, 4.08875, 4.2325, 4.3925, 4.42875, 4.75, 4.75, 4.91, 5.08875, 5.2675, 5.42875),
              "lucida.console"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.32125, 4.46375, 4.6075),
              "lucida.fax"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.7675),
              "lucida.handwriting"= c(1.25, 1.3925, 1.58875, 1.7325, 1.7325, 2.07125, 2.25, 2.25, 2.58875, 2.58875, 2.7675, 2.91, 3.05375, 3.25, 3.42875, 3.57125, 3.57125, 3.91, 3.94625, 4.08875, 4.2325, 4.3925, 4.6075, 4.75, 4.91, 4.91, 5.2675, 5.2675, 5.42875),
              "lucida.sans"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.6075),
              "lucida.sans.typewriter"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375),
              "lucida.sans.unicode"= c(1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125, 4.6075, 4.6075, 4.7675),
              "magneto"= c(1.17875, 1.32125, 1.625, 1.625, 1.80375, 2.1075, 2.1075, 2.25, 2.58875, 2.7325, 2.7325, 3.07125, 3.21375, 3.21375, 3.5, 3.69625, 3.69625, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925, 5.57125, 5.71375),
              "maiandra.gd"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "malgun.gothic"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.1425),
              "mangal"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "marlett"= c(1.80375, 1.96375, 2.25, 2.42875, 2.58875, 2.875, 3.07125, 3.21375, 3.5, 3.69625, 3.83875, 4.1425, 4.32125, 4.46375, 4.7675, 4.94625, 5.08875, 5.3925, 5.57125, 5.71375, 6.0175, 6.19625, 6.3575, 6.6425, 6.82125, 6.9825, 7.2675, 7.46375, 7.6075),
              "matura.mt.script.capitals"= c(1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 4.94625),
              "meiryo"= c(1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375, 4.7675, 4.7675, 4.94625),
              "meiryo.ui"= c(1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375, 4.7675, 4.7675, 4.94625),
              "microsoft.himalaya"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.42875),
              "microsoft.jhenghei"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "microsoft.jhenghei.ui"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125),
              "microsoft.new.tai.lue"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "microsoft.phagspa"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "microsoft.sans.serif"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "microsoft.tai.le"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "microsoft.uighur"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.58875),
              "microsoft.yahei"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "microsoft.yahei.ui"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "microsoft.yi.baiti"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "miriam"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "miriam.fixed"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "mistral"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "modern"= c(0.75, 0.75, 0.8925, 1.035, 1.035, 1.21375, 1.3575, 1.3575, 1.5175, 1.5175, 1.66, 1.83875, 1.83875, 2, 2.1425, 2.1425, 2.285, 2.46375, 2.46375, 2.625, 2.7675, 2.7675, 2.91, 2.91, 3.1075, 3.1075, 3.25, 3.3925, 3.3925),
              "modern.no..20"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "mongolian.baiti"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "monotype.corsiva"= c(0.785, 0.92875, 1.07125, 1.07125, 1.1075, 1.285, 1.42875, 1.42875, 1.625, 1.7675, 1.7675, 1.9825, 1.9825, 2.1425, 2.285, 2.32125, 2.46375, 2.46375, 2.66, 2.66, 2.83875, 2.9825, 2.9825, 3.125, 3.17875, 3.3575, 3.5, 3.5, 3.67875),
              "moolboran"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875),
              "ms.gothic"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ms.mincho"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ms.outlook"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ms.pgothic"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ms.pmincho"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ms.reference.sans.serif"= c(1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.7675),
              "ms.reference.specialty"= c(2.1075, 2.42875, 2.7325, 2.875, 3.07125, 3.5, 3.69625, 3.83875, 4.32125, 4.46375, 4.6075, 4.94625, 5.25, 5.3925, 5.71375, 5.875, 6.19625, 6.5, 6.6425, 6.82125, 7.2675, 7.46375, 7.6075, 8.08875, 8.2325, 8.375, 8.71375, 9, 9.16),
              "ms.sans.serif"= c(0.8925, 1.035, 1.035, 1.3575, 1.3575, 1.5175, 1.5175, 1.5175, 1.83875, 2, 2, 2.1425, 2.1425, 2.1425, 2.7675, 2.7675, 2.7675, 2.7675, 3.1075, 3.1075, 3.25, 3.25, 3.25, 3.25, 3.7325, 3.7325, 3.7325, 4.0175, 4.0175),
              "ms.serif"= c(0.75, 0.8925, 0.8925, 1.21375, 1.21375, 1.3575, 1.3575, 1.3575, 1.5175, 1.83875, 1.83875, 2.285, 2.285, 2.285, 2.46375, 2.46375, 2.46375, 2.46375, 2.7675, 2.7675, 2.7675, 3.875, 3.875, 3.875, 3.1075, 3.875, 3.875, 3.7325, 3.7325),
              "ms.ui.gothic"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "mt.extra"= c(1.80375, 1.96375, 2.25, 2.42875, 2.58875, 2.875, 3.07125, 3.21375, 3.5, 3.69625, 3.83875, 4.32125, 4.46375, 4.6075, 4.94625, 5.08875, 5.25, 5.57125, 5.71375, 5.875, 6.19625, 6.3575, 6.5, 6.82125, 6.9825, 7.125, 7.46375, 7.6075, 7.75),
              "mv.boli"= c(1.17875, 1.32125, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.3575, 3.3575, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.25),
              "narkisim"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.07125),
              "niagara.engraved"= c(0.57125, 0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075),
              "niagara.solid"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075),
              "nirmala.ui"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "nsimsun"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "nyala"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "ocr.a.extended"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.32125, 4.46375, 4.6075),
              "old.english.text.mt"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "onyx"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.17875, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.625, 1.96375, 1.96375, 1.96375, 2.1075, 2.1075, 2.1075),
              "palace.script.mt"= c(0.6075, 0.6075, 0.785, 0.785, 0.785, 0.92875, 0.92875, 1.07125, 1.1075, 1.285, 1.285, 1.285, 1.42875, 1.4825, 1.625, 1.625, 1.625, 1.7675, 1.80375, 1.9825, 1.9825, 2.1425, 2.1425, 2.1425, 2.32125, 2.32125, 2.46375, 2.46375, 2.46375),
              "palatino.linotype"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "papyrus"= c(1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.32125, 4.32125, 4.46375),
              "parchment"= c(0.3925, 0.3925, 0.3925, 0.3925, 0.57125, 0.57125, 0.57125, 0.57125, 0.71375, 0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1, 1, 1.17875, 1.17875, 1.17875, 1.17875, 1.32125, 1.32125, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825),
              "perpetua"= c(0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.21375),
              "perpetua.titling.mt"= c(1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 4.94625),
              "plantagenet.cherokee"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "playbill"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.58875),
              "pmingliu"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "pmingliu-extb"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5),
              "poor.richard"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425),
              "pristina"= c(0.71375, 1, 1.17875, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "raavi"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "rage.italic"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "rockwell"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875, 3.9825),
              "rockwell.condensed"= c(0.57125, 0.71375, 0.71375, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 2.875, 3.07125),
              "rockwell.extra.bold"= c(1, 1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.42875, 2.42875, 2.7325, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875),
              "rod"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "roman"= c(0.75, 0.75, 0.8925, 1.035, 1.035, 1.21375, 1.3575, 1.3575, 1.5175, 1.5175, 1.66, 1.83875, 1.83875, 2, 2.1425, 2.1425, 2.285, 2.46375, 2.46375, 2.625, 2.7675, 2.7675, 2.91, 2.91, 3.1075, 3.1075, 3.25, 3.3925, 3.3925),
              "sakkal.majalla"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "script"= c(0.6075, 0.6075, 0.75, 0.75, 0.8925, 1.035, 1.035, 1.035, 1.21375, 1.3575, 1.3575, 1.5175, 1.5175, 1.66, 1.83875, 1.83875, 1.83875, 2, 2.1425, 2.1425, 2.285, 2.285, 2.46375, 2.46375, 2.625, 2.625, 2.7675, 2.7675, 2.91),
              "script.mt.bold"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "segoe.print"= c(1.17875, 1.17875, 1.4825, 1.625, 1.625, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.875, 2.875, 3.07125, 3.3575, 3.3575, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925),
              "segoe.script"= c(1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 2.1075, 2.1075, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.3575, 3.3575, 3.5, 3.83875, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.25, 5.3925),
              "segoe.ui"= c(0.8575, 0.8575, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425),
              "segoe.ui.light"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.83875),
              "segoe.ui.semibold"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625, 3.83875, 3.9825, 3.9825),
              "segoe.ui.semilight"= c(1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425),
              "segoe.ui.symbol"= c(0.8575, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 3.9825),
              "shonar.bangla"= c(0.71375, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 2.875),
              "showcard.gothic"= c(1.17875, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.32125, 4.46375),
              "shruti"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875),
              "simhei"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "simplified.arabic"= c(0.8575, 0.8575, 1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425),
              "simplified.arabic.fixed"= c(1, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 3.9825, 4.1425, 4.32125, 4.46375, 4.46375),
              "simsun"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "simsun-extb"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "small.fonts"= c(0.8925, 0.8925, 1.21375, 1.5175, 1.5175, 1.5175, 1.83875, 1.83875, 1.83875, 2.285, 2.285, 2.285, 2.46375, 2.46375, 2.7675, 3.1075, 2.7675, 2.7675, 3.1075, 3.1075, 3.1075, 3.7325, 3.7325, 3.7325, 3.7325, 3.875, 3.875, 3.875, 3.875),
              "snap.itc"= c(1.4825, 1.625, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.25, 5.3925, 5.57125, 5.71375, 5.875, 6.0175, 6.3575, 6.5, 6.5),
              "sylfaen"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "symbol"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "tahoma"= c(1, 1, 1.32125, 1.32125, 1.4825, 1.625, 1.80375, 1.80375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.58875, 2.875, 2.875, 3.07125, 3.21375, 3.3575, 3.5, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.6075),
              "tempus.sans.itc"= c(1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.5, 3.69625, 3.83875, 3.9825, 4.1425, 4.1425, 4.46375, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875),
              "terminal"= c(1.21375, 1.21375, 1.21375, 1.83875, 1.83875, 1.5175, 1.5175, 1.5175, 1.5175, 2.58875, 2.58875, 2.58875, 2.58875, 2.58875, 3.7325, 3.7325, 3.7325, 3.7325, 3.1075, 3.1075, 3.1075, 4.46375, 4.46375, 4.46375, 4.46375, 4.46375, 5.6075, 5.6075, 5.6075),
              "times.new.roman"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "traditional.arabic"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "trebuchet.ms"= c(0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125),
              "tunga"= c(0.71375, 0.71375, 0.8575, 1, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575),
              "tw.cen.mt"= c(0.71375, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625, 3.83875),
              "tw.cen.mt.condensed"= c(0.57125, 0.71375, 0.8575, 0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.07125),
              "tw.cen.mt.condensed.extra.bold"= c(0.8575, 0.8575, 1, 1.17875, 1.17875, 1.32125, 1.4825, 1.4825, 1.625, 1.80375, 1.80375, 1.96375, 2.1075, 2.1075, 2.25, 2.42875, 2.42875, 2.58875, 2.7325, 2.7325, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5, 3.69625, 3.69625),
              "utsaah"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "vani"= c(1, 1.17875, 1.4825, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 2.875, 3.21375, 3.21375, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.08875),
              "verdana"= c(1.17875, 1.17875, 1.4825, 1.4825, 1.625, 1.96375, 1.96375, 2.1075, 2.25, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.83875, 3.83875, 4.1425, 4.1425, 4.32125, 4.6075, 4.6075, 4.7675, 4.94625, 5.08875, 5.25),
              "vijaya"= c(0.57125, 0.57125, 0.71375, 0.71375, 0.8575, 1, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.4825, 1.625, 1.625, 1.80375, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.42875, 2.58875, 2.58875, 2.7325, 2.7325, 2.875),
              "viner.hand.itc"= c(1.17875, 1.32125, 1.4825, 1.625, 1.80375, 1.96375, 2.1075, 2.1075, 2.42875, 2.42875, 2.58875, 2.7325, 2.875, 3.07125, 3.21375, 3.3575, 3.3575, 3.69625, 3.69625, 3.83875, 3.9825, 4.1425, 4.32125, 4.46375, 4.6075, 4.7675, 4.94625, 5.08875, 5.08875),
              "vivaldi"= c(0.785, 0.92875, 0.92875, 1.07125, 1.1075, 1.285, 1.42875, 1.42875, 1.625, 1.625, 1.7675, 1.94625, 1.9825, 1.9825, 2.1425, 2.285, 2.32125, 2.46375, 2.46375, 2.66, 2.83875, 2.83875, 2.83875, 2.9825, 3.17875, 3.17875, 3.3575, 3.3575, 3.5),
              "vladimir.script"= c(0.8575, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.07125, 3.21375, 3.3575, 3.3575, 3.5),
              "vrinda"= c(0.71375, 0.8575, 1, 1, 1.17875, 1.32125, 1.32125, 1.4825, 1.625, 1.625, 1.80375, 1.96375, 1.96375, 2.1075, 2.25, 2.25, 2.42875, 2.58875, 2.58875, 2.7325, 2.875, 2.875, 3.07125, 3.21375, 3.21375, 3.3575, 3.5, 3.5, 3.69625),
              "webdings"= c(1.80375, 1.96375, 2.25, 2.42875, 2.58875, 2.875, 3.07125, 3.21375, 3.5, 3.69625, 3.83875, 4.1425, 4.32125, 4.46375, 4.7675, 4.94625, 5.08875, 5.3925, 5.57125, 5.71375, 6.0175, 6.19625, 6.3575, 6.6425, 6.82125, 6.9825, 7.2675, 7.46375, 7.6075),
              "wide.latin"= c(2.25, 2.42875, 2.875, 3.07125, 3.21375, 3.69625, 3.83875, 3.9825, 4.32125, 4.6075, 4.7675, 5.08875, 5.3925, 5.57125, 5.875, 6.19625, 6.3575, 6.6425, 6.9825, 7.125, 7.46375, 7.6075, 7.8925, 8.2325, 8.375, 8.71375, 9, 9.16, 9.4825))


