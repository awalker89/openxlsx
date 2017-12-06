
#' @include class_definitions.R

Workbook$methods(initialize = function(creator = Sys.info()[["login"]]){
  
  if(length(creator) == 0)
    creator <- ""
  
  
  
  charts <<- list()
  isChartSheet <<- logical(0)
  
  colWidths <<- list()
  connections <<- NULL
  Content_Types <<- genBaseContent_Type()
  core <<- genBaseCore(creator)
  comments <<- list()
  
  
  drawings <<- list()
  drawings_rels <<- list()
  
  embeddings <<- NULL
  externalLinks <<- NULL
  externalLinksRels <<- NULL
  
  headFoot <<- NULL
  
  media <<- list()
  
  pivotTables <<- NULL
  pivotTables.xml.rels <<- NULL
  pivotDefinitions <<- NULL
  pivotRecords <<- NULL
  pivotDefinitionsRels <<- NULL
  
  queryTables <<- NULL
  rowHeights <<- list()
  
  slicers <<- NULL
  slicerCaches <<- NULL
  
  sheet_names <<- character(0)
  sheetOrder <<- integer(0)
  
  sharedStrings <<- list()
  attr(sharedStrings, "uniqueCount") <<- 0
  
  styles <<- genBaseStyleSheet()
  styleObjects <<- list()
  
  
  tables <<- NULL
  tables.xml.rels <<- NULL
  theme <<- NULL
  
  
  vbaProject <<- NULL
  vml <<- list()
  vml_rels <<- list()
  
  
  
  workbook <<- genBaseWorkbook()
  workbook.xml.rels <<- genBaseWorkbook.xml.rels()
  
  worksheets <<- list()
  worksheets_rels <<- list()
  
  
  
})







Workbook$methods(addWorksheet = function(sheetName
                                         , showGridLines = TRUE
                                         , tabColour = NULL
                                         , zoom = 100
                                         , oddHeader = NULL
                                         , oddFooter = NULL
                                         , evenHeader = NULL
                                         , evenFooter = NULL
                                         , firstHeader = NULL
                                         , firstFooter = NULL
                                         , visible = TRUE
                                         , paperSize = 9
                                         , orientation = 'portrait'
                                         , hdpi = 300
                                         , vdpi = 300){
  if (!missing(sheetName)) {
    if (grepl(pattern=":", x=sheetName)) stop("colon not allowed in sheet names in Excel")
  }	
  newSheetIndex = length(worksheets) + 1L
  
  if(newSheetIndex > 1){
    sheetId <- max(as.integer(regmatches(workbook$sheets, regexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE)))) + 1L
  }else{
    sheetId <- 1
  }
  
  
  ## fix visible value
  visible <- tolower(visible)
  if(visible == "true")
    visible <- "visible"
  
  if(visible == "false")
    visible <- "hidden"
  
  if(visible == "veryhidden")
    visible <- "veryHidden"
  
  
  ##  Add sheet to workbook.xml
  workbook$sheets <<- c(workbook$sheets, sprintf('<sheet name="%s" sheetId="%s" state="%s" r:id="rId%s"/>', sheetName, sheetId, visible, newSheetIndex))
  
  ## append to worksheets list
  worksheets <<- append(worksheets, WorkSheet$new( showGridLines = showGridLines
                                                   , tabSelected = newSheetIndex == 1
                                                   , tabColour = tabColour
                                                   , zoom = zoom
                                                   , oddHeader = oddHeader
                                                   , oddFooter = oddFooter
                                                   , evenHeader = evenHeader
                                                   , evenFooter = evenFooter
                                                   , firstHeader = firstHeader
                                                   , firstFooter = firstFooter
                                                   , paperSize = paperSize
                                                   , orientation = orientation
                                                   , hdpi = hdpi
                                                   , vdpi = vdpi))
  
  
  ## update content_tyes
  ## add a drawing.xml for the worksheet
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', newSheetIndex),
                      sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', newSheetIndex))
  
  ## Update xl/rels
  workbook.xml.rels <<- c(workbook.xml.rels,
                          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%s.xml"/>', newSheetIndex)
  )
  
  
  ## create sheet.rels to simplify id assignment
  worksheets_rels[[newSheetIndex]] <<- genBaseSheetRels(newSheetIndex)
  drawings_rels[[newSheetIndex]] <<- list()
  drawings[[newSheetIndex]] <<- list()
  
  vml_rels[[newSheetIndex]] <<- list()
  vml[[newSheetIndex]] <<- list()
  
  isChartSheet[[newSheetIndex]] <<- FALSE
  comments[[newSheetIndex]] <<- list()
  
  rowHeights[[newSheetIndex]] <<- list()
  colWidths[[newSheetIndex]] <<- list()
  
  sheetOrder <<- c(sheetOrder, as.integer(newSheetIndex))
  sheet_names <<- c(sheet_names, sheetName)
  
  invisible(newSheetIndex)
  
})


Workbook$methods(addChartSheet = function(sheetName, tabColour = NULL, zoom = 100){
  
  newSheetIndex <- length(worksheets) + 1L
  
  if(newSheetIndex > 1){
    sheetId <- max(as.integer(regmatches(workbook$sheets, regexpr('(?<=sheetId=")[0-9]+', workbook$sheets, perl = TRUE)))) + 1L
  }else{
    sheetId <- 1
  }
  
  ##  Add sheet to workbook.xml
  workbook$sheets <<- c(workbook$sheets, sprintf('<sheet name="%s" sheetId="%s" r:id="rId%s"/>', sheetName, sheetId, newSheetIndex))
  
  ## append to worksheets list
  worksheets <<- append(worksheets, ChartSheet$new(tabSelected = newSheetIndex == 1, tabColour = tabColour, zoom = zoom))
  sheet_names <<- c(sheet_names, sheetName)
  
  ## update content_tyes
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/chartsheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>', newSheetIndex))
  
  ## Update xl/rels
  workbook.xml.rels <<- c(workbook.xml.rels,
                          sprintf('<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet%s.xml"/>', newSheetIndex)
  )
  
  
  
  ## add a drawing.xml for the worksheet
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/drawings/drawing%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>', newSheetIndex))
  
  ## create sheet.rels to simplify id assignment
  worksheets_rels[[newSheetIndex]] <<- genBaseSheetRels(newSheetIndex)
  drawings_rels[[newSheetIndex]] <<- list()
  drawings[[newSheetIndex]] <<- list()
  
  isChartSheet[[newSheetIndex]] <<- TRUE
  
  rowHeights[[newSheetIndex]] <<- list()
  colWidths[[newSheetIndex]] <<- list()
  
  vml_rels[[newSheetIndex]] <<- list()
  vml[[newSheetIndex]] <<- list()
  
  sheetOrder <<- c(sheetOrder, newSheetIndex)
  
  invisible(newSheetIndex)
  
})



Workbook$methods(saveWorkbook = function(){
  
  ## temp directory to save XML files prior to compressing
  tmpDir <- file.path(tempfile(pattern = "workbookTemp_"))
  
  if(file.exists(tmpDir))
    unlink(tmpDir, recursive = TRUE, force = TRUE)
  
  success <- dir.create(path = tmpDir, recursive = TRUE)
  if(!success)
    stop(sprintf("Failed to create temporary directory '%s'", tmpDir))
  
  .self$preSaveCleanUp()
  
  nSheets <- length(worksheets)
  nThemes <- length(theme)
  nPivots <- length(pivotDefinitions)
  nSlicers <- length(slicers)
  nComments <- sum(sapply(comments, length) > 0)
  nVML <- sum(sapply(vml, length) > 0)
  
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
  
  
  ## will always have a theme
  xlthemeDir <- file.path(tmpDir, "xl", "theme")
  dir.create(path = xlthemeDir, recursive = TRUE)
  
  if(is.null(theme)){
    con <- file(file.path(xlthemeDir, "theme1.xml"), open = "wb")
    writeBin(charToRaw(genBaseTheme()), con)
    close(con)  
  }else{
    lapply(1:nThemes, function(i){
      con <- file(file.path(xlthemeDir, paste0("theme", i, ".xml")), open = "wb")
      writeBin(charToRaw(pxml(theme[[i]])), con)
      close(con)        
    })
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
  
  ## charts
  if(length(charts) > 0)
    file.copy(from = dirname(charts[1]), to = file.path(tmpDir, "xl"), recursive = TRUE)
  
  
  ## xl/comments.xml
  if(nComments > 0 | nVML > 0){
    for(i in 1:nSheets){
      if(length(comments[[i]]) > 0){
        
        fn <- sprintf("comments%s.xml", i)
        
        Content_Types <<- c(Content_Types, 
                            sprintf('<Override PartName="/xl/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>', fn))
        
        worksheets_rels[[i]] <<- unique(c(worksheets_rels[[i]],
                                          sprintf('<Relationship Id="rIdcomment" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../%s"/>', fn)))
        
        writeCommentXML(comment_list = comments[[i]], file_name = file.path(tmpDir, "xl", fn))
      }
    }
    
    .self$writeDrawingVML(xldrawingsDir)
    
  }
  
  if(length(embeddings) > 0){
    
    embeddingsDir <- file.path(tmpDir, "xl", "embeddings")
    dir.create(path = embeddingsDir, recursive = TRUE)
    for(fl in embeddings)
      file.copy(from = fl, to = embeddingsDir, overwrite = TRUE)
  }
  
  
  if(nPivots > 0){
    
    pivotTablesDir <- file.path(tmpDir, "xl", "pivotTables")
    dir.create(path = pivotTablesDir, recursive = TRUE)
    
    pivotTablesRelsDir <- file.path(tmpDir, "xl", "pivotTables", "_rels")
    dir.create(path = pivotTablesRelsDir, recursive = TRUE)
    
    pivotCacheDir <- file.path(tmpDir, "xl", "pivotCache")
    dir.create(path = pivotCacheDir, recursive = TRUE)
    
    pivotCacheRelsDir <- file.path(tmpDir, "xl", "pivotCache", "_rels")
    dir.create(path = pivotCacheRelsDir, recursive = TRUE)
    
    for(i in 1:length(pivotTables))
      file.copy(from = pivotTables[i], to =  file.path(pivotTablesDir, sprintf("pivotTable%s.xml", i)), overwrite = TRUE, copy.date = TRUE)
    
    for(i in 1:length(pivotDefinitions))
      file.copy(from = pivotDefinitions[i], to =  file.path(pivotCacheDir, sprintf("pivotCacheDefinition%s.xml", i)), overwrite = TRUE, copy.date = TRUE)    
    
    for(i in 1:length(pivotRecords))
      file.copy(from = pivotRecords[i], to =  file.path(pivotCacheDir, sprintf("pivotCacheRecords%s.xml", i)), overwrite = TRUE, copy.date = TRUE)
    
    for(i in 1:length(pivotDefinitionsRels))
      file.copy(from = pivotDefinitionsRels[i], to =  file.path(pivotCacheRelsDir, sprintf("pivotCacheDefinition%s.xml.rels", i)), overwrite = TRUE, copy.date = TRUE)
    
    for(i in 1:length(pivotTables.xml.rels))
      write_file(body = pivotTables.xml.rels[[i]], fl = file.path(pivotTablesRelsDir, sprintf("pivotTable%s.xml.rels", i)))
    
    
    
  }
  
  ## slicers
  if(nSlicers > 0){
    
    slicersDir <- file.path(tmpDir, "xl", "slicers")
    dir.create(path = slicersDir, recursive = TRUE)
    
    slicerCachesDir <- file.path(tmpDir, "xl", "slicerCaches")
    dir.create(path = slicerCachesDir, recursive = TRUE)
    
    for(i in 1:length(slicers)){
      if(nchar(slicers[i]) > 0)
        file.copy(from = slicers[i], to = file.path(slicersDir, sprintf("slicer%s.xml", i)))
    }
    
    
    
    for(i in 1:length(slicerCaches))
      write_file(body = slicerCaches[[i]], fl = file.path(slicerCachesDir, sprintf("slicerCache%s.xml", i)))
    
    
  }
  
  
  ## Write content
  
  ## write .rels
  write_file(head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
             body = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
                    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
             tail = '</Relationships>',
             fl = file.path(relsDir, ".rels"))
  
  
  ## write app.xml
  write_file(head = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
             body = '<Application>Microsoft Excel</Application>',
             tail = '</Properties>',
             fl = file.path(docPropsDir, "app.xml"))
  
  ## write core.xml
  write_file(head = '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
             body = pxml(core),
             tail = '</coreProperties>',
             fl = file.path(docPropsDir, "core.xml"))
  
  ## write workbook.xml.rels
  write_file(head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
             body = pxml(workbook.xml.rels),
             tail = '</Relationships>',
             fl = file.path(xlrelsDir, "workbook.xml.rels"))
  
  ## write tables
  if(length(unlist(tables, use.names = FALSE)) > 0){
    for(i in 1:length(unlist(tables, use.names = FALSE))){
      
      if(!grepl("openxlsx_deleted", attr(tables, "tableName")[i], fixed = TRUE)){
        write_file(body = pxml(unlist(tables, use.names = FALSE)[[i]]), fl = file.path(xlTablesDir, sprintf("table%s.xml", i+2)))
        if(tables.xml.rels[[i]] != "")
          write_file(body = tables.xml.rels[[i]], fl = file.path(xlTablesRelsDir, sprintf("table%s.xml.rels", i+2)))    
      }
      
    }
  }
  
  
  ## write query tables
  if(length(queryTables) > 0){
    xlqueryTablesDir <- file.path(tmpDir, "xl","queryTables")
    dir.create(path = xlqueryTablesDir, recursive = TRUE)
    
    for(i in 1:length(queryTables))
      write_file(body = queryTables[[i]], fl = file.path(xlqueryTablesDir, sprintf("queryTable%s.xml", i)))
  }
  
  ## connections
  if(length(connections) > 0)
    write_file(body = connections, fl = file.path(xlDir,"connections.xml"))
  
  ## externalLinks
  if(length(externalLinks)){
    externalLinksDir <- file.path(tmpDir, "xl","externalLinks")
    dir.create(path = externalLinksDir, recursive = TRUE)
    
    for(i in 1:length(externalLinks))
      write_file(body = externalLinks[[i]], fl = file.path(externalLinksDir, sprintf("externalLink%s.xml", i)))
  }
  
  ## externalLinks rels
  if(length(externalLinksRels)){
    externalLinksRelsDir <- file.path(tmpDir, "xl","externalLinks", "_rels")
    dir.create(path = externalLinksRelsDir, recursive = TRUE)
    
    for(i in 1:length(externalLinksRels))
      write_file(body = externalLinksRels[[i]], fl = file.path(externalLinksRelsDir, sprintf("externalLink%s.xml.rels", i)))
  } 
  
  # printerSettings
  printDir <- file.path(tmpDir, "xl", "printerSettings")
  dir.create(path = printDir, recursive = TRUE)
  for(i in 1:nSheets)
    writeLines(genPrinterSettings(), file.path(printDir, sprintf("printerSettings%s.bin", i)))
  
  ## media (copy file from origin to destination)
  for(x in media)
    file.copy(x, file.path(xlmediaDir, names(media)[which(media == x)]))
  
  ## VBA Macro
  if(!is.null(vbaProject))
    file.copy(vbaProject, xlDir)
  
  ## write worksheet, worksheet_rels, drawings, drawing_rels
  .self$writeSheetDataXML(xldrawingsDir, xldrawingsRelsDir, xlworksheetsDir, xlworksheetsRelsDir)
  
  ## write sharedStrings.xml
  ct <- Content_Types
  if(length(sharedStrings) > 0){
    write_file(head = sprintf('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%s" uniqueCount="%s">',   length(sharedStrings),   attr(sharedStrings, "uniqueCount")),
               body = paste(sharedStrings, collapse = ""),
               tail = "</sst>",
               fl = file.path(xlDir,"sharedStrings.xml"))
  }else{
    
    ## Remove relationship to sharedStrings
    ct <- ct[!grepl("sharedStrings", ct)]
  }
  
  if(nComments > 0)
    ct <- c(ct, '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>')
  
  ## write [Content_type]       
  write_file(head = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
             body = pxml(ct),
             tail = '</Types>',
             fl = file.path(tmpDir, "[Content_Types].xml"))
  
  
  styleXML <- styles
  styleXML$numFmts <- paste0(sprintf('<numFmts count="%s">', length(styles$numFmts)), pxml(styles$numFmts), '</numFmts>')
  styleXML$fonts <- paste0(sprintf('<fonts count="%s">', length(styles$fonts)), pxml(styles$fonts), '</fonts>')
  styleXML$fills <- paste0(sprintf('<fills count="%s">', length(styles$fills)), pxml(styles$fills), '</fills>')
  styleXML$borders <- paste0(sprintf('<borders count="%s">', length(styles$borders)), pxml(styles$borders), '</borders>')
  styleXML$cellStyleXfs <- c(sprintf('<cellStyleXfs count="%s">', length(styles$cellStyleXfs)), pxml(styles$cellStyleXfs), '</cellStyleXfs>')
  styleXML$cellXfs <- paste0(sprintf('<cellXfs count="%s">', length(styles$cellXfs)), pxml(styles$cellXfs), '</cellXfs>')
  styleXML$cellStyles <- paste0(sprintf('<cellStyles count="%s">', length(styles$cellStyles)), pxml(styles$cellStyles), '</cellStyles>')
  styleXML$dxfs <- ifelse(length(styles$dxfs) == 0, '<dxfs count="0"/>', paste0(sprintf('<dxfs count="%s">', length(styles$dxfs)), paste(unlist(styles$dxfs), collapse = ""), '</dxfs>'))
  
  ## write styles.xml
  write_file(head = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">',
             body = pxml(styleXML),
             tail = '</styleSheet>',
             fl = file.path(xlDir,"styles.xml"))
  
  ## write workbook.xml
  workbookXML <- workbook
  workbookXML$sheets <- paste0("<sheets>", pxml(workbookXML$sheets), "</sheets>")
  if(length(workbookXML$definedNames) > 0)
    workbookXML$definedNames <- paste0("<definedNames>", pxml(workbookXML$definedNames), "</definedNames>")
  
  
  write_file(head = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
             body = pxml(workbookXML),
             tail = '</workbook>',
             fl = file.path(xlDir,"workbook.xml"))
  workbook$sheets <<- workbook$sheets[order(sheetOrder)] ## Need to reset sheet order to allow multiple savings
  
  ## compress to xlsx
  wd <- getwd()
  tmpFile <- basename(tempfile(fileext = ifelse(is.null(vbaProject), ".xlsx", ".xlsm")))
  on.exit(expr = setwd(wd), add = TRUE)
  
  ## zip it
  setwd(dir = tmpDir)
  cl <- ifelse(!is.null(getOption("openxlsx.compresssionLevel")), getOption("openxlsx.compresssionLevel"), getOption("openxlsx.compresssionevel", 6))
  zip::zip(zipfile = tmpFile, files = list.files(tmpDir, full.names = FALSE, recursive = TRUE, include.dirs = FALSE, all.files=TRUE), compression_level = cl)  
  
  ## reset styles - maintain any changes to base font
  baseFont <- styles$fonts[[1]]
  styles <<- genBaseStyleSheet(styles$dxfs, tableStyles = styles$tableStyles, extLst = styles$extLst)
  styles$fonts[[1]] <<- baseFont
  
  
  return(file.path(tmpDir, tmpFile))
  
})



Workbook$methods(updateSharedStrings = function(uNewStr){
  
  ## Function will return named list of references to new strings  
  uStr <- uNewStr[which(!uNewStr %in% sharedStrings)]
  uCount <- attr(sharedStrings, "uniqueCount")
  sharedStrings <<- append(sharedStrings, uStr)
  
  attr(sharedStrings, "uniqueCount") <<- uCount + length(uStr)
  
})



Workbook$methods(validateSheet = function(sheetName){
  
  if(!is.numeric(sheetName))
    sheetName <- replaceIllegalCharacters(sheetName)
  
  if(is.null(sheet_names))
    stop("Workbook does not contain any worksheets.", call.=FALSE)
  
  if(is.numeric(sheetName)){
    if(sheetName > length(sheet_names))
      stop(sprintf("This Workbook only has %s sheets.", length(sheet_names)), call.=FALSE)
    
    return(sheetName)
    
  }else if(!sheetName %in% sheet_names){
    stop(sprintf("Sheet '%s' does not exist.", sheetName), call.=FALSE)
  }
  
  return(which(sheet_names == sheetName))
  
})



Workbook$methods(getSheetName = function(sheetIndex){
  
  if(any(length(sheet_names) < sheetIndex))
    stop(sprintf("Workbook only contains %s sheet(s).", length(sheet_names)))
  
  sheet_names[sheetIndex]
  
})



Workbook$methods(buildTable = function(sheet, colNames, ref, showColNames, tableStyle, tableName, withFilter
                                       , totalsRowCount = 0
                                       , showFirstColumn = 0
                                       , showLastColumn = 0
                                       , showRowStripes = 1
                                       , showColumnStripes = 0){
  
  ## id will start at 3 and drawing will always be 1, printer Settings at 2 (printer settings has been removed)
  id <- as.character(length(tables) + 3L)
  sheet <- validateSheet(sheet)
  
  ## build table XML and save to tables field
  table <- sprintf('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%s" name="%s" displayName="%s" ref="%s" totalsRowCount="%s"',
                   id,
                   tableName,
                   tableName,
                   ref,
                   as.integer(totalsRowCount))
  
  nms <- names(tables)
  tSheets <- attr(tables, "sheet")
  tNames <- attr(tables, "tableName") 
  
  tableStyleXML <- sprintf('<tableStyleInfo name="%s" showFirstColumn="%s" showLastColumn="%s" showRowStripes="%s" showColumnStripes="%s"/>', 
                           tableStyle, as.integer(showFirstColumn), as.integer(showLastColumn), as.integer(showRowStripes), as.integer(showColumnStripes))
  
  
  tables <<- c(tables, build_table_xml(table = table, tableStyleXML = tableStyleXML, ref = ref, colNames = gsub("\n|\r", "_x000a_", colNames), showColNames = showColNames, withFilter = withFilter))
  names(tables) <<- c(nms, ref)
  attr(tables, "sheet") <<- c(tSheets, sheet)
  attr(tables, "tableName") <<- c(tNames, tableName)
  
  worksheets[[sheet]]$tableParts <<- append(worksheets[[sheet]]$tableParts, sprintf('<tablePart r:id="rId%s"/>', id))
  attr(worksheets[[sheet]]$tableParts, "tableName") <<- c(tNames[tSheets == sheet & !grepl("openxlsx_deleted", tNames, fixed = TRUE)], tableName)
  
  
  
  ## update Content_Types
  Content_Types <<- c(Content_Types, sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', id))
  
  ## create a table.xml.rels
  tables.xml.rels <<- append(tables.xml.rels, "")
  
  ## update worksheets_rels
  worksheets_rels[[sheet]] <<- c(worksheets_rels[[sheet]],
                                 sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',  id,  id))
  
  
})









Workbook$methods(writeDrawingVML = function(dir){
  
  for(i in 1:length(comments)){
    
    id <- 1025
    
    cd <- unlist(lapply(comments[[i]],"[[", "clientData"))
    nComments <- length(cd)
    
    ## write head
    if(nComments > 0 | length(vml[[i]]) > 0){
      
      write(x = paste('<xml xmlns:v="urn:schemas-microsoft-com:vml"
                    xmlns:o="urn:schemas-microsoft-com:office:office"
                    xmlns:x="urn:schemas-microsoft-com:office:excel">
                    <o:shapelayout v:ext="edit">
                    <o:idmap v:ext="edit" data="1"/>
                    </o:shapelayout>
                    <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
                    path="m,l,21600r21600,l21600,xe">
                    <v:stroke joinstyle="miter"/>
                    <v:path gradientshapeok="t" o:connecttype="rect"/>
                    </v:shapetype>'), file = file.path(dir, sprintf("vmlDrawing%s.vml", i)))
      
    }
    
    if(nComments > 0){
      for(j in 1:nComments){
        id <- id + 1L
        write(x = genBaseShapeVML(cd[j], id), file = file.path(dir, sprintf("vmlDrawing%s.vml", i)), append = TRUE)
      }
    }
    
    if(length(vml[[i]]) > 0)
      write(x = vml[[i]], file = file.path(dir, sprintf("vmlDrawing%s.vml", i)), append = TRUE)
    
    if(nComments > 0 | length(vml[[i]]) > 0){
      write(x = '</xml>', file = file.path(dir, sprintf("vmlDrawing%s.vml", i)), append = TRUE)
      worksheets[[i]]$legacyDrawing <<- '<legacyDrawing r:id="rIdvml"/>'
    }
    
  }
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
    fontId <- which(styles$fonts == fontNode) - 1L
    
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
  if(!is.null(style$fill)){
    fillNode <- .self$createFillNode(style)
    if(!is.null(fillNode)){
      fillId <- which(styles$fills == fillNode) - 1L
      
      if(length(fillId) == 0){      
        fillId <- length(styles$fills)
        styles$fills <<- c(styles$fills, fillNode)
      }
      xfNode$fillId <- fillId
      xfNode <- append(xfNode, list("applyFill" = "1"))
    }
  }
  
  ## Border
  if(any(!is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom, style$borderDiagonal)))){
    
    borderNode <- .self$createBorderNode(style)
    borderId <- which(styles$borders == borderNode) - 1L
    
    if(length(borderId) == 0){
      borderId <- length(styles$borders)
      styles$borders <<- c(styles$borders, borderNode)
    }
    
    xfNode$borderId <- borderId
    xfNode <- append(xfNode, list("applyBorder" = "1"))
  }
  
  
  # if(!is.null(style$xfId))
  # xfNode$xfId <- style$xfId
  
  
  ## Alignment
  if(!is.null(style$halign) | !is.null(style$valign) | !is.null(style$wrapText) | !is.null(style$textRotation) | !is.null(style$indent)){
    
    attrs <- list()
    alignNode <- "<alignment"
    
    if(!is.null(style$textRotation))
      alignNode <- paste(alignNode, sprintf('textRotation="%s"', style$textRotation))
    
    if(!is.null(style$halign))
      alignNode <- paste(alignNode, sprintf('horizontal="%s"', style$halign))
    
    if(!is.null(style$valign))
      alignNode <- paste(alignNode, sprintf('vertical="%s"', style$valign))
    
    if(!is.null(style$indent))
      alignNode <- paste(alignNode, sprintf('indent="%s"', style$indent))
    
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





Workbook$methods(updateCellStyles = function(){
  
  
  flag <- TRUE
  for(style in cellStyleObjects){
    
    ## Updates styles.xml
    xfNode <- list(numFmtId = 0,
                   fontId = 0,
                   fillId = 0,
                   borderId = 0)
    
    
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
    if(!is.null(style$fill)){
      fillNode <- .self$createFillNode(style)
      if(!is.null(fillNode)){
        fillId <- which(styles$fills == fillNode) - 1L
        
        if(length(fillId) == 0){      
          fillId <- length(styles$fills)
          styles$fills <<- c(styles$fills, fillNode)
        }
        xfNode$fillId <- fillId
        xfNode <- append(xfNode, list("applyFill" = "1"))
      }
    }
    
    ## Border
    if(any(!is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom, style$borderDiagonal)))){
      
      borderNode <- .self$createBorderNode(style)
      borderId <- which(styles$borders == borderNode) - 1L
      
      if(length(borderId) == 0){
        borderId <- length(styles$borders)
        styles$borders <<- c(styles$borders, borderNode)
      }
      
      xfNode$borderId <- borderId
      xfNode <- append(xfNode, list("applyBorder" = "1"))
    }
    
    xfNode <- paste0("<xf ", paste(paste0(names(xfNode), '="',xfNode, '"'), collapse = " "), "/>")
    
    if(flag){
      styles$cellStyleXfs <<- xfNode
      flag <- FALSE
    }else{
      styles$cellStyleXfs <<- c(styles$cellStyleXfs, xfNode)
    }
    
    
  }
  
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
    
    if(length(style$fontColour) > 1){
      
      fontNode <- paste0(fontNode, sprintf('<color %s/>', 
                                           paste(sapply(1:length(style$fontColour), function(i) sprintf('%s="%s"', names(style$fontColour)[i], style$fontColour[i])), collapse = " ")))
      
    }else{
      fontNode <- paste0(fontNode, sprintf('<color %s="%s"/>', names(style$fontColour), style$fontColour))
    }
    
    
    
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
  
  borderNode <- "<border"
  
  if(style$borderDiagonalUp)
    borderNode <- paste(borderNode, 'diagonalUp="1"', sep = " ")
  
  if(style$borderDiagonalDown)
    borderNode <- paste(borderNode, 'diagonalDown="1"', sep = " ")
  
  borderNode <- paste0(borderNode, ">")
    
  if(!is.null(style$borderLeft))
    borderNode <- paste0(borderNode, sprintf('<left style="%s">', style$borderLeft), sprintf('<color %s="%s"/>', names(style$borderLeftColour), style$borderLeftColour), '</left>')
  
  if(!is.null(style$borderRight))
    borderNode <- paste0(borderNode, sprintf('<right style="%s">', style$borderRight), sprintf('<color %s="%s"/>', names(style$borderRightColour), style$borderRightColour), '</right>')
  
  if(!is.null(style$borderTop))
    borderNode <- paste0(borderNode, sprintf('<top style="%s">', style$borderTop), sprintf('<color %s="%s"/>', names(style$borderTopColour), style$borderTopColour), '</top>')
  
  if(!is.null(style$borderBottom))
    borderNode <- paste0(borderNode, sprintf('<bottom style="%s">', style$borderBottom), sprintf('<color %s="%s"/>', names(style$borderBottomColour), style$borderBottomColour), '</bottom>')
  
  if(!is.null(style$borderDiagonal))
    borderNode <- paste0(borderNode, sprintf('<diagonal style="%s">', style$borderDiagonal), sprintf('<color %s="%s"/>', names(style$borderDiagonalColour), style$borderDiagonalColour), '</diagonal>')
  
  paste0(borderNode, "</border>")
  
})


Workbook$methods(createFillNode = function(style, patternType = "solid"){
  
  fill <- style$fill
  
  ## gradientFill
  if(any(grepl("gradientFill", fill))){
    
    fillNode <- fill #paste0("<fill>", fill, "</fill>")
    
  }else if(!is.null(fill$fillFg) | !is.null(fill$fillBg)){
    
    fillNode <- paste0('<fill>', sprintf('<patternFill patternType="%s">', patternType))
    
    if(!is.null(fill$fillFg))
      fillNode <- paste0(fillNode, sprintf('<fgColor %s/>', paste(paste0(names(fill$fillFg), '="', fill$fillFg, '"'), collapse = " ")))
    
    if(!is.null(fill$fillBg))
      fillNode <- paste0(fillNode, sprintf('<bgColor %s/>', paste(paste0(names(fill$fillBg), '="', fill$fillBg, '"'), collapse = " ")))
    
    fillNode <- paste0(fillNode, "</patternFill></fill>")
    
  }else{
    return(NULL)
  }
  
  return(fillNode)
  
})







Workbook$methods(setSheetName = function(sheet, newSheetName){
  
  if(newSheetName %in% sheet_names)
    stop(sprintf("Sheet %s already exists!", newSheetName))  
  
  sheet <- validateSheet(sheet)
  
  oldName <- sheet_names[[sheet]]
  sheet_names[[sheet]] <<- newSheetName
  
  ## Rename in workbook
  sheetId <- regmatches(workbook$sheets[[sheet]], regexpr('(?<=sheetId=")[0-9]+', workbook$sheets[[sheet]], perl = TRUE))
  rId <- regmatches(workbook$sheets[[sheet]], regexpr('(?<= r:id="rId)[0-9]+', workbook$sheets[[sheet]], perl = TRUE))
  workbook$sheets[[sheet]] <<- sprintf('<sheet name="%s" sheetId="%s" r:id="rId%s"/>', newSheetName, sheetId, rId)
  
  ## rename styleObjects sheet component
  if(length(styleObjects) > 0){
    
    styleObjects <<- lapply(styleObjects, function(x){
      
      if(x$sheet == oldName)
        x$sheet <- newSheetName
      
      return(x)      
      
    })
    
  }
  
  ## rename defined names
  if(length(workbook$definedNames) > 0){
    
    belongTo <- getDefinedNamesSheet(workbook$definedNames)
    toChange <- belongTo == oldName
    if(any(toChange)){
      
      newSheetName <- sprintf("'%s'", newSheetName)
      tmp <- gsub(oldName, newSheetName, workbook$definedName[toChange], fixed = TRUE)
      tmp <- gsub("'+", "'", tmp)
      workbook$definedNames[toChange] <<- tmp
      
    }
    
    
  }
  
})


Workbook$methods(writeSheetDataXML = function(xldrawingsDir, xldrawingsRelsDir, xlworksheetsDir, xlworksheetsRelsDir){
  
  ## write worksheets  
  nSheets <- length(worksheets)
  
  for(i in 1:nSheets){
    
    ## Write drawing i (will always exist) skip those that are empty
    if(any(drawings[[i]] != "")){
      write_file(head = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                 body = pxml(drawings[[i]]),
                 tail = '</xdr:wsDr>',
                 fl =  file.path(xldrawingsDir, paste0("drawing", i,".xml")))
      
      write_file(head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                 body = pxml(drawings_rels[[i]]),
                 tail = '</Relationships>',
                 fl =  file.path(xldrawingsRelsDir, paste0("drawing", i,".xml.rels")))
    }else{
      worksheets[[i]]$drawing <<- character(0)
    }
    
    ## vml drawing
    if(length(vml_rels[[i]]) > 0)
      file.copy(from = vml_rels[[i]], to = file.path(xldrawingsRelsDir, paste0("vmlDrawing", i,".vml.rels")))
    
    
    
    if(isChartSheet[i]){
      
      chartSheetDir <- file.path(dirname(xlworksheetsDir), "chartsheets")
      chartSheetRelsDir <- file.path(dirname(xlworksheetsDir), "chartsheets", "_rels")
      
      if(!file.exists(chartSheetDir)){
        dir.create(chartSheetDir, recursive = TRUE)
        dir.create(chartSheetRelsDir, recursive = TRUE)
      }
      
      write_file(body = worksheets[[i]]$get_prior_sheet_data(), fl = file.path(chartSheetDir, paste0("sheet", i,".xml")))
      
      write_file(head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">', 
                 body = pxml(worksheets_rels[[i]]),
                 tail = '</Relationships>',
                 fl = file.path(chartSheetRelsDir, sprintf("sheet%s.xml.rels", i)))
      
    }else{
      
      
      ## Write worksheets
      ws <- worksheets[[i]]
      hasHL <- ifelse(length(worksheets[[i]]$hyperlinks) > 0, TRUE, FALSE)
      
      ## reorder sheet data
      worksheets[[i]]$order_sheetdata()
      
      
      prior <- ws$get_prior_sheet_data()
      post <- ws$get_post_sheet_data()
      
      worksheets[[i]]$sheet_data$style_id <<- as.character(worksheets[[i]]$sheet_data$style_id)
      
      if(length(rowHeights[[i]]) == 0){
        
        write_worksheet_xml(prior = prior,
                            post = post,
                            sheet_data = ws$sheet_data,
                            R_fileName = file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)))
        
      }else{
        
        ## row heights will always be in order and all row heights are given rows in preSaveCleanup  
        write_worksheet_xml_2(prior = prior,
                              post = post,
                              sheet_data = ws$sheet_data,
                              row_heights =  unlist(rowHeights[[i]]),
                              R_fileName = file.path(xlworksheetsDir, sprintf("sheet%s.xml", i)))
        
      }
      
      worksheets[[i]]$sheet_data$style_id <<- integer(0)
      
      
      ## write worksheet rels
      if(length(worksheets_rels[[i]]) > 0 ){
        
        ws_rels <- worksheets_rels[[i]]
        if(hasHL){
          h_inds <- paste0(1:length(worksheets[[i]]$hyperlinks), "h")
          ws_rels <- c(ws_rels, unlist(lapply(1:length(h_inds), function(j) worksheets[[i]]$hyperlinks[[j]]$to_target_xml(h_inds[j]))))
        }
        
        ## Check if any tables were deleted - remove these from rels
        if(length(tables) > 0){
          table_inds <- which(grepl("tables/table[0-9].xml", ws_rels))
          
          if(length(table_inds) > 0){
            
            ids <- regmatches(ws_rels[table_inds], regexpr('(?<=Relationship Id=")[0-9A-Za-z]+', ws_rels[table_inds], perl = TRUE))
            inds <- as.integer(gsub("[^0-9]", "", ids, perl = TRUE)) - 2L
            table_nms <- attr(tables, "tableName")[inds]
            is_deleted <- grepl("openxlsx_deleted", table_nms, fixed = TRUE)
            if(any(is_deleted))
              ws_rels <- ws_rels[-table_inds[is_deleted]]
            
          }
        }
        
        
        
        write_file(head = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">', 
                   body = pxml(ws_rels),
                   tail = '</Relationships>',
                   fl = file.path(xlworksheetsRelsDir, sprintf("sheet%s.xml.rels", i))
        )
        
      }
      
    }## end of isChartSheet[i]
    
  } ## end of loop through 1:nSheets
  
  invisible(0)
  
})





Workbook$methods(setRowHeights = function(sheet, rows, heights){
  
  sheet = validateSheet(sheet)
  
  ## remove any conflicting heights
  flag <- names(rowHeights[[sheet]]) %in% rows
  if(any(flag))
    rowHeights[[sheet]] <<- rowHeights[[sheet]][!flag]
  
  nms <- c(names(rowHeights[[sheet]]), rows)
  allRowHeights <- unlist(c(rowHeights[[sheet]], heights))
  names(allRowHeights) <- nms
  
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
  
  # Remove vml element
  # Remove vml_rels element
  
  # Remove rowHeights element
  # Remove styleObjects on sheet
  # Remove last sheet element from workbook
  # Remove last sheet element from workbook.xml.rels
  # Remove element from worksheets
  # Remove element from worksheets_rels
  # Remove hyperlinks
  # Reduce calcChain i attributes & remove calcs on sheet
  # Remove sheet from sheetOrder
  # Remove queryTable references from workbook$definedNames to worksheet
  # remove tables
  
  sheet <- validateSheet(sheet)
  sheetNames <- sheet_names
  nSheets <- length(unlist(sheetNames, use.names = FALSE))
  sheetName <- sheetNames[[sheet]]
  
  colWidths[[sheet]] <<- NULL
  sheet_names <<- sheet_names[-sheet]
  
  ## remove last drawings(sheet).xml from Content_Types
  Content_Types <<- Content_Types[!grepl(sprintf("drawing%s.xml", nSheets), Content_Types)]
  
  ## remove highest sheet
  Content_Types <<- Content_Types[!grepl(sprintf("sheet%s.xml", nSheets), Content_Types)]
  
  drawings[[sheet]] <<- NULL
  drawings_rels[[sheet]] <<- NULL
  
  vml[[sheet]] <<- NULL
  vml_rels[[sheet]] <<- NULL
  
  rowHeights[[sheet]] <<- NULL
  comments[[sheet]] <<- NULL
  isChartSheet <<- isChartSheet[-sheet]
  
  ## sheetOrder
  toRemove <- which(sheetOrder == sheet)
  sheetOrder[sheetOrder > sheet] <<- sheetOrder[sheetOrder > sheet] - 1L
  sheetOrder <<- sheetOrder[-toRemove]
  
  
  ## remove styleObjects
  if(length(styleObjects) > 0)
    styleObjects <<- styleObjects[unlist(lapply(styleObjects, "[[", "sheet"), use.names = FALSE) != sheetName]
  
  ## Need to remove reference from workbook.xml.rels to pivotCache
  removeRels <- worksheets_rels[[sheet]][grepl("pivotTables", worksheets_rels[[sheet]])]
  if(length(removeRels) > 0){
    
    ## sheet rels links to a pivotTable file, the corresponding pivotTable_rels file links to the cacheDefn which is listing in workbook.xml.rels
    ## remove reference to this file from the workbook.xml.rels
    fileNo <- as.integer(unlist(regmatches(removeRels, gregexpr('(?<=pivotTable)[0-9]+(?=\\.xml)', removeRels, perl = TRUE))))
    toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", fileNo), collapse = "|")    
    
    fileNo <- which(grepl(toRemove, pivotTables.xml.rels))
    toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", fileNo), collapse = "|")
    
    ## remove reference to file from workbook.xml.res
    workbook.xml.rels <<- workbook.xml.rels[!grepl(toRemove, workbook.xml.rels)]
    
  }
  
  ## As above for slicers
  ## Need to remove reference from workbook.xml.rels to pivotCache
  removeRels <- grepl("slicers", worksheets_rels[[sheet]])
  if(any(removeRels))
    workbook.xml.rels <<- workbook.xml.rels[!grepl(sprintf("(slicerCache%s\\.xml)", sheet), workbook.xml.rels)]
  
  ## wont't remove tables and then won't need to reassign table r:id's but will rename them!
  worksheets[[sheet]] <<- NULL
  worksheets_rels[[sheet]] <<- NULL
  
  if(length(tables) > 0){
    
    tableSheets <- attr(tables, "sheet")
    tableNames <- attr(tables, "tableName")
    
    inds <- tableSheets %in% sheet & !grepl("openxlsx_deleted", attr(tables, "tableName"), fixed = TRUE)
    tableSheets[tableSheets > sheet] <- tableSheets[tableSheets > sheet] - 1L
    
    ## Need to flag a table as deleted
    if(any(inds)){
      tableSheets[inds] <- 0
      tableNames[inds] <- paste0(tableNames[inds], "_openxlsx_deleted")
    }
    attr(tables, "tableName") <<- tableNames
    attr(tables, "sheet") <<- tableSheets
  }
  
  
  ## drawing will always be the first relationship and printerSettings second
  if(nSheets > 1){
    for(i in 1:(nSheets-1L))
      worksheets_rels[[i]][1:3] <<- genBaseSheetRels(i)
  }else{
    worksheets_rels <<- list()
  }
  
  
  ## remove sheet
  sn <- unlist(lapply(workbook$sheets, function(x) regmatches(x, regexpr('(?<= name=")[^"]+', x, perl = TRUE))))
  workbook$sheets <<- workbook$sheets[!sn %in% sheetName]
  
  ## Reset rIds
  if(nSheets > 1){
    for(i in (sheet+1L):nSheets)
      workbook$sheets <<- gsub(paste0("rId", i), paste0("rId", i-1L), workbook$sheets, fixed = TRUE)
  }else{
    workbook$sheets <<- NULL
  }
  
  ## Can remove highest sheet
  workbook.xml.rels <<- workbook.xml.rels[!grepl(sprintf("sheet%s.xml", nSheets), workbook.xml.rels)]
  
  ## definedNames
  if(length(workbook$definedNames) > 0){
    belongTo <- getDefinedNamesSheet(workbook$definedNames)
    workbook$definedNames <<- workbook$definedNames[!belongTo %in% sheetName]
  }
  
  invisible(1)
  
})


Workbook$methods(addDXFS = function(style){
  
  dxf <- '<dxf>'
  dxf <- paste0(dxf, createFontNode(style))
  fillNode <- NULL
  
  if(!is.null(style$fill$fillFg) | !is.null(style$fill$fillBg))
    dxf <- paste0(dxf, createFillNode(style))
  
  if(any(!is.null(c(style$borderLeft, style$borderRight, style$borderTop, style$borderBottom, style$borderDiagonal))))
    dxf <- paste0(dxf, createBorderNode(style))
  
  dxf <- paste(dxf, "</dxf>")
  if(dxf %in% styles$dxfs)
    return(which(styles$dxfs == dxf) - 1L)
  
  dxfId <- length(styles$dxfs)
  styles$dxfs <<- c(styles$dxfs, dxf)
  
  return(dxfId)
})



Workbook$methods(dataValidation = function(sheet, startRow, endRow, startCol, endCol, type, operator, value, allowBlank, showInputMsg, showErrorMsg){
  
  sheet = validateSheet(sheet)
  sqref <- paste(getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol))), collapse = ":")
  
  header <- sprintf('<dataValidation type="%s" operator="%s" allowBlank="%s" showInputMessage="%s" showErrorMessage="%s" sqref="%s">',
                    type, operator, allowBlank, showInputMsg, showErrorMsg, sqref)
  
  
  if(type == "date"){
    
    origin <- 25569L
    if(grepl('date1904="1"|date1904="true"', paste(unlist(workbook), collapse = ""), ignore.case = TRUE))
      origin <- 24107L
    
    value <- as.integer(value) + origin
    
  }
  
  if(type == "time"){
    
    origin <- 25569L
    if(grepl('date1904="1"|date1904="true"', paste(unlist(workbook), collapse = ""), ignore.case = TRUE))
      origin <- 24107L
    
    t <- format(value[1], "%z")
    offSet <- suppressWarnings(ifelse(substr(t,1,1) == "+", 1L, -1L) * (as.integer(substr(t,2,3)) + as.integer(substr(t,4,5)) / 60) / 24)
    if(is.na(offSet)) offSet[i] <- 0
    
    value <- as.numeric(as.POSIXct(value)) / 86400 + origin + offSet
    
  }    
  
  form <- sapply(1:length(value), function(i) sprintf("<formula%s>%s</formula%s>", i, value[i], i))
  worksheets[[sheet]]$dataValidations <<- c(worksheets[[sheet]]$dataValidations, paste0(header, paste(form, collapse = ""), "</dataValidation>"))
  
  invisible(0)
  
})



Workbook$methods(dataValidation_list = function(sheet, startRow, endRow, startCol, endCol, value, allowBlank, showInputMsg, showErrorMsg){
  
  sheet = validateSheet(sheet)
  sqref <- paste(getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol))), collapse = ":")
  
  header <- '<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
			        <x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
  
  data_val <- sprintf('<x14:dataValidation type="list" allowBlank="%s" showInputMessage="%s" showErrorMessage="%s">',
                      allowBlank, showInputMsg, showErrorMsg, sqref)
  
  formula <- sprintf("<x14:formula1><xm:f>%s</xm:f></x14:formula1>", value)
  sqref <- sprintf('<xm:sqref>%s</xm:sqref>', sqref)
  
  xml <- paste0(header, data_val, formula, sqref, '</x14:dataValidation></x14:dataValidations></ext>')
  
  worksheets[[sheet]]$extLst <<- c(worksheets[[sheet]]$extLst, xml)
  
  invisible(0)
  
})






Workbook$methods(conditionalFormatting = function(sheet, startRow, endRow, startCol, endCol, dxfId, formula, type, values, params){
  
  sheet = validateSheet(sheet)
  sqref <- paste(getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol))), collapse = ":")
  
  
  
  ## Increment priority of conditional formatting rule
  if(length(worksheets[[sheet]]$conditionalFormatting) > 0){
    
    for(i in length(worksheets[[sheet]]$conditionalFormatting):1){
      
      priority <- regmatches(worksheets[[sheet]]$conditionalFormatting[[i]], regexpr('(?<=priority=")[0-9]+', worksheets[[sheet]]$conditionalFormatting[[i]], perl = TRUE))
      priority_new <- as.integer(priority) + 1L
      
      priority_pattern <- sprintf('priority="%s"', priority)
      priority_new <- sprintf('priority="%s"', priority_new)
      
      ## now replace
      worksheets[[sheet]]$conditionalFormatting[[i]] <<- gsub(priority_pattern, priority_new, worksheets[[sheet]]$conditionalFormatting[[i]], fixed = TRUE)
      
    }
    
  }
  
  nms <- c(names(worksheets[[sheet]]$conditionalFormatting), sqref)
  
  if(type == "colorScale"){
    
    ## formula contains the colours
    ## values contains numerics or is NULL
    ## dxfId is ignored
    
    if(is.null(values)){
      
      if(length(formula) == 2L){
        cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale>
                             <cfvo type="min"/><cfvo type="max"/>
                             <color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>', formula[[1]], formula[[2]])
      }else{
        cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale>
                             <cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>
                             <color rgb="%s"/><color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>', formula[[1]], formula[[2]], formula[[3]])
      }
      
    }else{
      
      if(length(formula) == 2L){
        cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale>
                            <cfvo type="num" val="%s"/><cfvo type="num" val="%s"/>
                            <color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>', values[[1]], values[[2]], formula[[1]], formula[[2]])
      }else{
        cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale>
                            <cfvo type="num" val="%s"/><cfvo type="num" val="%s"/><cfvo type="num" val="%s"/>
                            <color rgb="%s"/><color rgb="%s"/><color rgb="%s"/>
                           </colorScale></cfRule>', values[[1]], values[[2]], values[[3]], formula[[1]], formula[[2]], formula[[3]])
      }
      
    }
    
    
  }else if(type == "dataBar"){
    
    # forumula is a vector of colours of length 1 or 2
    # values is NULL or a numeric vector of equal length as formula
    
    if(length(formula) == 2L){
      negColour <- formula[[1]]
      posColour <- formula[[2]]
    }else{
      posColour <- formula
      negColour <- "FFFF0000"
    }
    
    guid <- paste0("F7189283-14F7-4DE0-9601-54DE9DB", 40000L + length(worksheets[[sheet]]$extLst))
    
    showValue <- 1
    if("showValue" %in% names(params))
      showValue <- as.integer(params$showValue)
    
    gradient <- 1
    if("gradient" %in% names(params))
      gradient <- as.integer(params$gradient)
    
    border <- 1
    if("border" %in% names(params))
      border <- as.integer(params$border)
    
    if(is.null(values)){
      
      
      cfRule <- sprintf('<cfRule type="dataBar" priority="1"><dataBar showValue="%s">
                          <cfvo type="min"/><cfvo type="max"/>
                          <color rgb="%s"/>
                          </dataBar>
                          <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{%s}</x14:id></ext>
                        </extLst></cfRule>', showValue, posColour, guid)
      
    }else{
      
      cfRule <- sprintf('<cfRule type="dataBar" priority="1"><dataBar showValue="%s">
                            <cfvo type="num" val="%s"/><cfvo type="num" val="%s"/>
                            <color rgb="%s"/>
                            </dataBar>
                            <extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                            <x14:id>{%s}</x14:id></ext></extLst></cfRule>', showValue, values[[1]], values[[2]], posColour, guid)
      
    }
    
    worksheets[[sheet]]$extLst <<- c(worksheets[[sheet]]$extLst, gen_databar_extlst(guid = guid,
                                                                                    sqref = sqref,
                                                                                    posColour = posColour,
                                                                                    negColour = negColour,
                                                                                    values = values, 
                                                                                    border = border, 
                                                                                    gradient = gradient))
    
  }else if(type == "expression"){
    
    cfRule <- sprintf('<cfRule type="expression" dxfId="%s" priority="1"><formula>%s</formula></cfRule>', dxfId, formula)
    
    
  }else if(type == "duplicatedValues"){
    
    cfRule <- sprintf('<cfRule type="duplicateValues" dxfId="%s" priority="1"/>', dxfId)
    
    
    
  }else if(type == "containsText"){
    
    cfRule <- sprintf('<cfRule type="containsText" dxfId="%s" priority="1" operator="containsText" text="%s">
                        	<formula>NOT(ISERROR(SEARCH("%s", %s)))</formula>
                       </cfRule>', dxfId, values, values, unlist(strsplit(sqref, split = ":"))[[1]])
    
  }else if(type == "between"){
    
    cfRule <- sprintf('<cfRule type="cellIs" dxfId="%s" priority="1" operator="between"><formula>%s</formula><formula>%s</formula></cfRule>', dxfId, formula[1], formula[2])
    
    
  }
  
  worksheets[[sheet]]$conditionalFormatting <<- append(worksheets[[sheet]]$conditionalFormatting, cfRule)              
  
  names(worksheets[[sheet]]$conditionalFormatting) <<- nms
  
  invisible(0)
  
})




Workbook$methods(mergeCells = function(sheet, startRow, endRow, startCol, endCol) {
  
  sheet <- validateSheet(sheetName = sheet)
  
  sqref <- getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol)))
  exMerges <- regmatches(worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))
  
  if(!is.null(exMerges)){
    
    comps <- lapply(exMerges, function(rectCoords) unlist(strsplit(rectCoords, split = ":")))  
    exMergedCells <- build_cell_merges(comps = comps)
    newMerge <- unlist(build_cell_merges(comps = list(sqref)))
    
    ## Error if merge intersects    
    mergeIntersections <- sapply(exMergedCells, function(x) any(x %in% newMerge))
    if(any(mergeIntersections))
      stop(sprintf("Merge intersects with existing merged cells: \n\t\t%s.\nRemove existing merge first.", paste(exMerges[mergeIntersections], collapse = "\n\t\t")))
    
  }
  
  worksheets[[sheet]]$mergeCells <<- c(worksheets[[sheet]]$mergeCells, sprintf('<mergeCell ref="%s"/>', paste(sqref, collapse = ":")))
  
  
})



Workbook$methods(removeCellMerge = function(sheet, startRow, endRow, startCol, endCol){
  
  sheet = validateSheet(sheet)
  
  sqref <- getCellRefs(data.frame("x" = c(startRow, endRow), "y" = c(startCol, endCol)))
  exMerges <- regmatches(worksheets[[sheet]]$mergeCells, regexpr("[A-Z0-9]+:[A-Z0-9]+", worksheets[[sheet]]$mergeCells))
  
  if(!is.null(exMerges)){
    
    comps <- lapply(exMerges, function(x) unlist(strsplit(x, split = ":")))  
    exMergedCells <- build_cell_merges(comps = comps)
    newMerge <- unlist(build_cell_merges(comps = list(sqref)))
    
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
    
    if(firstActiveRow == 1 & firstActiveCol == 1) ## nothing to do
      return(NULL)
    
    if(firstActiveRow > 1 & firstActiveCol == 1){
      attrs <- sprintf('ySplit="%s"', firstActiveRow - 1L)
      activePane <- "bottomLeft"
    }
    
    if(firstActiveRow == 1 & firstActiveCol > 1){
      attrs <- sprintf('xSplit="%s"', firstActiveCol - 1L)
      activePane <- "topRight"
    }
    
    if(firstActiveRow > 1 & firstActiveCol > 1){
      attrs <- sprintf('ySplit="%s" xSplit="%s"', firstActiveRow - 1L, firstActiveCol - 1L)
      activePane <- "bottomRight"
    }
    
    topLeftCell <- getCellRefs(data.frame(firstActiveRow, firstActiveCol))
    
    paneNode <- sprintf('<pane %s topLeftCell="%s" activePane="%s" state="frozen"/><selection pane="%s"/>', 
                        paste(attrs, collapse = " "), topLeftCell, activePane, activePane)
    
  } 
  
  worksheets[[sheet]]$freezePane <<- paneNode
  
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
  ## vmlDrawing to have rId 
  
  sheetRIds <- as.integer(unlist(regmatches(workbook$sheets, gregexpr('(?<=r:id="rId)[0-9]+', workbook$sheets, perl = TRUE))))
  
  nSheets <- length(sheetRIds)
  nExtRefs <- length(externalLinks)
  nPivots <- length(pivotDefinitions)
  
  ## add a worksheet if none added
  if(nSheets == 0){
    warning("Workbook does not contain any worksheets. A worksheet will be added.", call. = FALSE)
    .self$addWorksheet("Sheet 1")
    nSheets <- 1L
  }
  
  ## get index of each child element for ordering
  sheetInds <- which(grepl("(worksheets|chartsheets)/sheet[0-9]+\\.xml", workbook.xml.rels))
  stylesInd <- which(grepl("styles\\.xml", workbook.xml.rels))
  themeInd <- which(grepl("theme/theme[0-9]+.xml", workbook.xml.rels))
  connectionsInd <- which(grepl("connections.xml", workbook.xml.rels))
  extRefInds <- which(grepl("externalLinks/externalLink[0-9]+.xml", workbook.xml.rels))
  sharedStringsInd <- which(grepl("sharedStrings.xml", workbook.xml.rels))
  tableInds <- which(grepl("table[0-9]+.xml", workbook.xml.rels))
  
  
  ## Reordering of workbook.xml.rels
  ## don't want to re-assign rIds for pivot tables or slicer caches
  pivotNode <- workbook.xml.rels[grepl("pivotCache/pivotCacheDefinition[0-9].xml", workbook.xml.rels)]
  slicerNode <- workbook.xml.rels[which(grepl("slicerCache[0-9]+.xml", workbook.xml.rels))]
  
  ## Reorder children of workbook.xml.rels
  workbook.xml.rels <<- workbook.xml.rels[c(sheetInds, extRefInds, themeInd, connectionsInd, stylesInd, sharedStringsInd, tableInds)]
  
  ## Re assign rIds to children of workbook.xml.rels
  workbook.xml.rels <<- unlist(lapply(1:length(workbook.xml.rels), function(i) {
    gsub('(?<=Relationship Id="rId)[0-9]+', i, workbook.xml.rels[[i]], perl = TRUE)
  }))
  
  workbook.xml.rels <<- c(workbook.xml.rels, pivotNode, slicerNode)
  
  
  
  if(!is.null(vbaProject))
    workbook.xml.rels <<- c(workbook.xml.rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>', 1L + length(workbook.xml.rels)))
  
  ## Reassign rId to workbook sheet elements, (order sheets by sheetId first)
  workbook$sheets <<- unlist(lapply(1:length(workbook$sheets), function(i) {
    gsub('(?<= r:id="rId)[0-9]+', i, workbook$sheets[[i]], perl = TRUE)
  }))
  
  ## re-order worksheets if need to
  if(any(sheetOrder != 1:nSheets))
    workbook$sheets <<- workbook$sheets[sheetOrder]
  
  
  
  ## re-assign tabSelected
  state <- rep.int("visible", nSheets)
  state[grepl("hidden", workbook$sheets)] <- "hidden"
  visible_sheet_index <- which(state %in% "visible")[[1]]
  
  workbook$bookViews <<- sprintf('<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="13125" windowHeight="6105" firstSheet="%s" activeTab="%s"/></bookViews>',
                                 visible_sheet_index-1L,
                                 visible_sheet_index-1L)
  
  worksheets[[visible_sheet_index]]$sheetViews <<- sub('( tabSelected="0")|( tabSelected="false")', ' tabSelected="1"', worksheets[[visible_sheet_index]]$sheetViews, ignore.case = TRUE)
  if(nSheets > 1){
    for(i in (1:nSheets)[!(1:nSheets) %in% visible_sheet_index])
      worksheets[[i]]$sheetViews <<- sub(' tabSelected="(1|true|false|0)"', ' tabSelected="0"', worksheets[[i]]$sheetViews, ignore.case = TRUE)
  }
  
  
  
  
  
  if(length(workbook$definedNames) > 0){
    
    sheetNames <- sheet_names[sheetOrder]
    
    belongTo <- getDefinedNamesSheet(workbook$definedNames)
    
    ## sheetNames is in re-ordered order (order it will be displayed)
    newId <- match(belongTo, sheetNames) - 1L
    oldId <- as.numeric(regmatches(workbook$definedNames, regexpr('(?<= localSheetId=")[0-9]+', workbook$definedNames, perl = TRUE)))
    
    for(i in 1:length(workbook$definedNames)){
      if(!is.na(newId[i]))
        workbook$definedNames[[i]] <<- gsub(sprintf('localSheetId=\"%s\"', oldId[i]), sprintf('localSheetId=\"%s\"', newId[i]), workbook$definedNames[[i]], fixed = TRUE)
    }
  }
  
  
  
  
  ## update workbook r:id to match reordered workbook.xml.rels externalLink element
  if(length(extRefInds) > 0){
    newInds <- as.integer(1:length(extRefInds) + length(sheetInds))
    workbook$externalReferences <<- paste0("<externalReferences>",
                                           paste0(sprintf('<externalReference r:id=\"rId%s\"/>', newInds), collapse = ""),
                                           "</externalReferences>")
  }
  
  ## styles
  numFmtIds <- 50000L
  for(i in which(!isChartSheet))
    worksheets[[i]]$sheet_data$style_id <<- rep.int(x = as.integer(NA), times = worksheets[[i]]$sheet_data$n_elements)
  
  
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
      sheet <- which(sheet_names == x$sheet)
      sId <- .self$updateStyles(this.sty) ## this creates the XML for styles.XML
      
      cells_to_style <- paste(x$rows, x$cols, sep = ",")
      existing_cells <- paste(worksheets[[sheet]]$sheet_data$rows, worksheets[[sheet]]$sheet_data$cols, sep = ",")
      
      ## In here we create any style_ids that don't yet exist in sheet_data
      worksheets[[sheet]]$sheet_data$style_id[existing_cells %in% cells_to_style] <<- sId
      
      
      new_cells_to_append <- which(!cells_to_style %in% existing_cells)
      if(length(new_cells_to_append) > 0){
        
        worksheets[[sheet]]$sheet_data$style_id <<- c(worksheets[[sheet]]$sheet_data$style_id, rep.int(x = sId, times = length(new_cells_to_append)))
        
        worksheets[[sheet]]$sheet_data$rows <<- c(worksheets[[sheet]]$sheet_data$rows, x$rows[new_cells_to_append])
        worksheets[[sheet]]$sheet_data$cols <<- c(worksheets[[sheet]]$sheet_data$cols, x$cols[new_cells_to_append])
        worksheets[[sheet]]$sheet_data$t <<- c(worksheets[[sheet]]$sheet_data$t, rep(as.integer(NA), length(new_cells_to_append)))
        worksheets[[sheet]]$sheet_data$v <<- c(worksheets[[sheet]]$sheet_data$v, rep(as.character(NA), length(new_cells_to_append)))
        worksheets[[sheet]]$sheet_data$f <<- c(worksheets[[sheet]]$sheet_data$f, rep(as.character(NA), length(new_cells_to_append)))
        worksheets[[sheet]]$sheet_data$data_count <<- worksheets[[sheet]]$sheet_data$data_count + 1L
        
        worksheets[[sheet]]$sheet_data$n_elements <<- as.integer(length(worksheets[[sheet]]$sheet_data$rows))
        
        
      }
    }
  }
  
  
  ## Make sure all rowHeights have rows, if not append them!
  for(i in 1:length(worksheets)){
    
    if(length(rowHeights[[i]]) > 0){
      
      rh <- as.integer(names(rowHeights[[i]]))
      missing_rows <- rh[!rh %in% worksheets[[i]]$sheet_data$rows]
      n <- length(missing_rows)
      
      if(n > 0){
        
        worksheets[[i]]$sheet_data$style_id <<- c(worksheets[[i]]$sheet_data$style_id, rep.int(as.integer(NA), times = n))
        
        worksheets[[i]]$sheet_data$rows <<- c(worksheets[[i]]$sheet_data$rows, missing_rows)
        worksheets[[i]]$sheet_data$cols <<- c(worksheets[[i]]$sheet_data$cols, rep.int(as.integer(NA), times = n))
        
        worksheets[[i]]$sheet_data$t <<- c(worksheets[[i]]$sheet_data$t, rep(as.integer(NA), times = n))
        worksheets[[i]]$sheet_data$v <<- c(worksheets[[i]]$sheet_data$v, rep(as.character(NA), times = n))
        worksheets[[i]]$sheet_data$f <<- c(worksheets[[i]]$sheet_data$f, rep(as.character(NA), times = n))
        worksheets[[i]]$sheet_data$data_count <<- worksheets[[i]]$sheet_data$data_count + 1L
        
        worksheets[[i]]$sheet_data$n_elements <<- as.integer(length(worksheets[[i]]$sheet_data$rows))
        
      }
    }
    
    ## write colwidth XML   
    if(length(colWidths[[i]]) > 0)
      invisible(.self$setColWidths(i))
    
  }
  
})



Workbook$methods(addStyle = function(sheet, style, rows, cols, stack){
  
  sheet <- sheet_names[[sheet]]
  
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
    keepStyle <- rep(TRUE, nStyles) 
    for(i in 1:nStyles){
      
      if(sheet == styleObjects[[i]]$sheet){
        
        ## Now check rows and cols intersect
        ## toRemove are the elements that the new style doesn't apply to, we remove these from the style object as it
        ## is copied, merged with the new style and given the new data points
        
        ex_row_cols <- paste(styleObjects[[i]]$rows, styleObjects[[i]]$cols, sep = "-")
        new_row_cols <- paste(rows, cols, sep = "-")
        
        
        ## mergeInds are the intersection of the two styles that will need to merge
        mergeInds <- which(new_row_cols %in% ex_row_cols)
        
        ## newInds are inds that don't exist in the current - this cumulates until the end to see if any are new
        newInds <- newInds[!newInds %in% mergeInds]
        
        
        ## If the new style does not merge 
        if(length(mergeInds) > 0){
          
          to_remove_from_this_style_object <- which(ex_row_cols %in% new_row_cols)
          
          ## the new style intersects with this styleObjects[[i]], we need to remove the intersecting rows and
          ## columns from styleObjects[[i]]
          if(length(to_remove_from_this_style_object) > 0){
            
            ## remove these from style object
            styleObjects[[i]]$rows <<- styleObjects[[i]]$rows[-to_remove_from_this_style_object]
            styleObjects[[i]]$cols <<- styleObjects[[i]]$cols[-to_remove_from_this_style_object]
            
            if(length(styleObjects[[i]]$rows) == 0 | length(styleObjects[[i]]$cols) == 0)
              keepStyle[i] <- FALSE ## this style applies to no rows or columns anymore
            
            
          }
          
          ## append style object for intersecting cells
          
          ## we are appending a new style
          keepStyle <- c(keepStyle, TRUE) ## keepStyle is used to remove styles that apply to 0 rows OR 0 columns
          
          ## Merge Style and append to styleObjects
          styleObjects <<- append(styleObjects, list(list(style = mergeStyle(styleObjects[[i]]$style, newStyle = style),
                                                          sheet = sheet,
                                                          rows = rows[mergeInds],
                                                          cols = cols[mergeInds])))        
          
        } 
        
        
      } ## if sheet == styleObjects[[i]]$sheet
      
      
    } ## End of loop through styles
    
    ## remove any styles that no longer have any affect
    if(!all(keepStyle))
      styleObjects <<- styleObjects[keepStyle]
    
    ## append style object for non-intersecting cells
    if(length(newInds) > 0){
      
      styleObjects <<- append(styleObjects, list(list(style = style,
                                                      sheet =  sheet,
                                                      rows = rows[newInds],
                                                      cols = cols[newInds])))
      
    }
    
    
  }else{ ## else we are not stacking
    
    styleObjects <<- append(styleObjects, list(list(style = style,
                                                    sheet =  sheet,
                                                    rows = rows,
                                                    cols = cols)))
    
    
  } ## End if(length(styleObjects) > 0) else if(stack) {}
  
  
  
  
})



Workbook$methods(createNamedRegion = function(ref1, ref2, name, sheet, localSheetId = NULL){
  
  name <- replaceIllegalCharacters(name)
  
  if(is.null(localSheetId)){
    workbook$definedNames <<- c(workbook$definedNames, 
                                sprintf('<definedName name="%s">\'%s\'!%s:%s</definedName>', name, sheet, ref1, ref2 )
    )
  }else{
    workbook$definedNames <<- c(workbook$definedNames, 
                                sprintf('<definedName name="%s" localSheetId="%s">\'%s\'!%s:%s</definedName>', name, localSheetId, sheet, ref1, ref2)
    )
  }
  
})


Workbook$methods(validate_table_name = function(tableName){
  
  tableName <- tolower(tableName) ## Excel forces named regions to lowercase
  
  if(nchar(tableName) > 255)
    stop("tableName must be less than 255 characters.")
  
  if(grepl("$", tableName, fixed = TRUE))
    stop("'$' character cannot exist in a tableName")
  
  if(grepl(" ", tableName, fixed = TRUE))
    stop("spaces cannot exist in a table name")
  
  # if(!grepl("^[A-Za-z_]", tableName, perl = TRUE))
  #   stop("tableName must begin with a letter or an underscore")
  
  if(grepl("R[0-9]+C[0-9]+", tableName, perl = TRUE, ignore.case = TRUE))
    stop("tableName cannot be the same as a cell reference, such as R1C1")
  
  if(grepl('^[A-Z]{1,3}[0-9]+$', tableName, ignore.case = TRUE))
    stop("tableName cannot be the same as a cell reference")
  
  if(tableName %in% attr(tables, "tableName"))
    stop(sprintf("Table with name '%s' already exists!", tableName))
  
  return(tableName)
  
})


Workbook$methods(check_overwrite_tables = function(sheet
                                                   , new_rows
                                                   , new_cols
                                                   , error_msg = "Cannot overwrite existing table with another table."
                                                   , check_table_header_only = FALSE){
  
  
  
  ## check not overwriting another table
  if(length(tables) > 0){
    
    tableSheets <- attr(tables, "sheet")
    sheetNo <- validateSheet(sheet)
    
    to_check <- which(tableSheets %in% sheetNo & !grepl("openxlsx_deleted", attr(tables, "tableName"), fixed = TRUE))
    
    if(length(to_check) > 0){ ## only look at tables on this sheet
      
      exTable <- tables[to_check]
      
      rows <- lapply(names(exTable), function(rectCoords) as.numeric(unlist(regmatches(rectCoords, gregexpr("[0-9]+", rectCoords)))))
      cols <- lapply(names(exTable), function(rectCoords) convertFromExcelRef(unlist(regmatches(rectCoords, gregexpr("[A-Z]+", rectCoords)))))
      
      if(check_table_header_only)
        rows <- lapply(rows, function(x) c(x[1], x[1]))
      
      
      ## loop through existing tables checking if any over lap with new table
      for(i in 1:length(exTable)){
        
        existing_cols <- cols[[i]]
        existing_rows <- rows[[i]]
        
        if((min(new_cols) <= max(existing_cols)) & (max(new_cols) >= min(existing_cols)) & (min(new_rows) <= max(existing_rows)) & (max(new_rows) >= min(existing_rows)))
          stop(error_msg)
        
      }
    } ## end if(sheet %in% tableSheets) 
  } ## end (length(tables) > 0)
  
  invisible(0)  
  
})




Workbook$methods(show = function(){
  
  
  exSheets <- sheet_names
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
        cols <- names(colWidths[[i]])
        widths <- unname(colWidths[[i]])
        
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









## TO BE DEPRECATED
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
    
  }else if(length(formula) == 2L){
    
    cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="%s"/><color rgb="%s"/></colorScale></cfRule>', formula[[1]], formula[[2]])
    
  }else{
    
    cfRule <- sprintf('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="%s"/><color rgb="%s"/><color rgb="%s"/></colorScale></cfRule>', formula[[1]], formula[[2]], formula[[3]])
  }
  
  worksheets[[sheet]]$conditionalFormatting <<- append(worksheets[[sheet]]$conditionalFormatting, cfRule)              
  
  names(worksheets[[sheet]]$conditionalFormatting) <<- nms
  
  invisible(0)
  
})






Workbook$methods(loadStyles = function(stylesXML){
  
  ## Build style objects from the styles XML
  stylesTxt <- readLines(stylesXML, warn = FALSE, encoding = "UTF-8")
  stylesTxt <- removeHeadTag(stylesTxt)
  
  ## Indexed colours
  vals <- getNodes(xml = stylesTxt, tagIn = "<indexedColors>")
  if(length(vals) > 0)
    styles$indexedColors <<- paste0("<colors>", vals, "</colors>")
  
  ## dxf (don't need these, I don't think)
  dxf <- getNodes(xml = stylesTxt, tagIn = "<dxfs")
  if(length(dxf) > 0){
    dxf <- getNodes(xml = dxf[[1]], tagIn = "<dxf>")
    if(length(dxf) > 0)
      styles$dxfs <<- dxf
  }
  
  tableStyles <- getNodes(xml = stylesTxt, tagIn = "<tableStyles")
  if(length(tableStyles) > 0)
    styles$tableStyles <<- paste0(tableStyles, ">")
  
  extLst <- getNodes(xml = stylesTxt, tagIn = "<extLst>")
  if(length(extLst) > 0)
    styles$extLst <<- extLst
  
  
  ## Number formats
  numFmts <- getChildlessNode(xml = stylesTxt, tag = "<numFmt ")
  numFmtFlag <- FALSE
  if(length(numFmts) > 0){
    
    numFmtsIds <- sapply(numFmts, getAttr, tag = 'numFmtId="', USE.NAMES = FALSE)
    formatCodes <- sapply(numFmts, getAttr, tag = 'formatCode="', USE.NAMES = FALSE)
    numFmts <-lapply(1:length(numFmts), function(i) list("numFmtId"= numFmtsIds[[i]], "formatCode"=formatCodes[[i]]))
    numFmtFlag <- TRUE
    
  }
  
  ## fonts will maintain, sz, color, name, family scheme
  fonts <- getNodes(xml = stylesTxt, tagIn = "<font>")
  styles$fonts[[1]] <<- fonts[[1]]
  fonts <- buildFontList(fonts)       
  
  fills <- getNodes(xml = stylesTxt, tagIn = "<fill>")
  fills <- buildFillList(fills)
  
  borders <- getOpenClosedNode(stylesTxt, "<borders ", "</borders>")
  borders <- substr(borders, start =  regexpr("<border>", borders)[1], stop = regexpr("</borders>", borders) - 1L)
  borders <- getNodes(xml = borders, tagIn = "<border")
  borders <- sapply(borders, buildBorder, USE.NAMES = FALSE)
  

  ## ------------------------------ build styleObjects ------------------------------ ##
  
  cellXfs <- getNodes(xml = stylesTxt, tagIn = "<cellXfs")
  
  xf <- getChildlessNode(xml = cellXfs, tag = "<xf ")
  xfAttrs <- regmatches(xf, gregexpr('[a-zA-Z]+=".*?"', xf))
  xfNames <- lapply(xfAttrs, function(xfAttrs) regmatches(xfAttrs, regexpr('[a-zA-Z]+(?=\\=".*?")', xfAttrs, perl = TRUE)))
  xfVals <- lapply(xfAttrs, function(xfAttrs) regmatches(xfAttrs, regexpr('(?<=").*?(?=")', xfAttrs, perl = TRUE)))
  
  for(i in 1:length(xf))
    names(xfVals[[i]]) <- xfNames[[i]]
  
  styleObjects_tmp <- list()
  flag <- FALSE
  for(s in xfVals){
    
    style <- createStyle()
    if(any(s != "0")){
      
      if("fontId" %in% names(s)){
        if(s[["fontId"]] != "0"){
          thisFont <- fonts[[(as.integer(s[["fontId"]])+1)]]
          
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
      }
      
      if("numFmtId" %in% names(s)){
        if(s[["numFmtId"]] != "0"){
          if(as.integer(s[["numFmtId"]]) < 164){
            style$numFmt <- list(numFmtId = s[["numFmtId"]])
          }else if(numFmtFlag){
            style$numFmt <- numFmts[[which(s[["numFmtId"]] == numFmtsIds)[1]]]
          }
        }
      }
      
      ## Border
      if("borderId" %in% names(s)){
        if(s[["borderId"]] != "0"){# & "applyBorder" %in% names(s)){
          
          border_ind <- as.integer(s[["borderId"]]) + 1L
          if(border_ind <= length(borders)){
            thisBorder <- borders[[border_ind]]
            
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
            
            if("borderDiagonal" %in% names(thisBorder)){
              style$borderDiagonal    <- thisBorder$borderDiagonal
              style$borderDiagonalColour <- thisBorder$borderDiagonalColour
            }
            
            if("borderDiagonalUp" %in% names(thisBorder))
              style$borderDiagonalUp    <- thisBorder$borderDiagonalUp
            
            if("borderDiagonalDown" %in% names(thisBorder))
              style$borderDiagonalDown    <- thisBorder$borderDiagonalDown
            
            
          }
        }
      }
      
      ## alignment
      # applyAlignment <- "applyAlignment" %in% names(s)
      if("horizontal" %in% names(s))# & applyAlignment)
        style$halign <- s[["horizontal"]]
      
      if("vertical" %in% names(s))
        style$valign <- s[["vertical"]]  
      
      if("indent" %in% names(s))
        style$indent <- s[["indent"]]  
      
      if("textRotation" %in% names(s))
        style$textRotation <- s[["textRotation"]]
      
      ## wrap text
      if("wrapText" %in% names(s)){
        if(s[["wrapText"]] %in% c("1", "true"))
          style$wrapText <- TRUE
      }
      
      if("fillId" %in% names(s)){
        if(s[["fillId"]] != "0"){
          
          fillId <- as.integer(s[["fillId"]]) + 1L
          
          if("fgColor" %in% names(fills[[fillId]])){
            
            tmpFg <- fills[[fillId]]$fgColor
            tmpBg <- fills[[fillId]]$bgColor
            
            if(!is.null(tmpFg))
              style$fill$fillFg <- tmpFg
            
            if(!is.null(tmpFg))
              style$fill$fillBg <- tmpBg
          }else{
            style$fill <- fills[[fillId]]
          }
          
        }
      }
      
      
      if("xfId" %in% names(s)){
        if(s[["xfId"]] != "0"){
          style$xfId <- s[["xfId"]]
        }
      }
      
      
    } ## end if !all(s == "0")
    
    ## we need to skip the first one as this is used as the base style
    if(flag)
      styleObjects_tmp <- append(styleObjects_tmp , list(style))
    
    flag <- TRUE
    
  }  ## end of for loop through styles s in ...
  
  
  ## ------------------------------ build styleObjects Complete ------------------------------ ##
  
  
  return(styleObjects_tmp)
  
  
})



