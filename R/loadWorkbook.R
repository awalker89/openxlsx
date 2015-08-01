


#' @name loadWorkbook 
#' @title Load an exisiting .xlsx file
#' @author Alexander Walker
#' @param file A path to an existing .xlsx or .xlsm file
#' @param xlsxFile alias for file
#' @description  loadWorkbook returns a workbook object conserving styles and 
#' formatting of the original .xlsx file. 
#' @return Workbook object. 
#' @export
#' @seealso \code{\link{removeWorksheet}}
#' @examples
#' ## load existing workbook from package folder
#' wb <- loadWorkbook(file = system.file("loadExample.xlsx", package= "openxlsx"))
#' names(wb)  #list worksheets
#' wb ## view object
#' ## Add a worksheet
#' addWorksheet(wb, "A new worksheet")
#' 
#' ## Save workbook
#' saveWorkbook(wb, "loadExample.xlsx", overwrite = TRUE)
loadWorkbook <- function(file, xlsxFile = NULL){
  
  if(!is.null(xlsxFile))
    file <- xlsxFile
  
  if(!file.exists(file))
    stop("File does not exist.")
  
  wb <- createWorkbook()
  
  ## create temp dir
  xmlDir <- file.path(tempdir(),  paste0(tempfile(tmpdir = ""), "_openxlsx_loadWorkbook"))
  
  ## Unzip files to temp directory
  xmlFiles <- unzip(file, exdir = xmlDir)
  
  .relsXML           <- xmlFiles[grepl("_rels/.rels$", xmlFiles, perl = TRUE)]
  drawingsXML        <- xmlFiles[grepl("drawings/drawing[0-9]+.xml$", xmlFiles, perl = TRUE)]
  worksheetsXML      <- xmlFiles[grepl("/worksheets/sheet[0-9]", xmlFiles, perl = TRUE)]
  appXML             <- xmlFiles[grepl("app.xml$", xmlFiles, perl = TRUE)]
  coreXML            <- xmlFiles[grepl("core.xml$", xmlFiles, perl = TRUE)]
  workbookXML        <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  stylesXML          <- xmlFiles[grepl("styles.xml$", xmlFiles, perl = TRUE)]
  sharedStringsXML   <- xmlFiles[grepl("sharedStrings.xml$", xmlFiles, perl = TRUE)]
  themeXML           <- xmlFiles[grepl("theme[0-9]+.xml$", xmlFiles, perl = TRUE)]
  drawingRelsXML     <- xmlFiles[grepl("drawing[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  sheetRelsXML       <- xmlFiles[grepl("sheet[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  media              <- xmlFiles[grepl("image[0-9]+.[a-z]+$", xmlFiles, perl = TRUE)]
  vmlDrawingXML      <- xmlFiles[grepl("drawings/vmlDrawing[0-9]+\\.vml$", xmlFiles, perl = TRUE)]
  commentsXML        <- xmlFiles[grepl("xl/comments[0-9]+\\.xml", xmlFiles, perl = TRUE)]
  
  
  charts             <- xmlFiles[grepl("xl/charts/.*xml$", xmlFiles, perl = TRUE)]
  chartsRels         <- xmlFiles[grepl("xl/charts/_rels", xmlFiles, perl = TRUE)]
  chartSheetsXML     <- xmlFiles[grepl("xl/chartsheets/sheet[0-9]+\\.xml", xmlFiles, perl = TRUE)]
  
  tablesXML          <- xmlFiles[grepl("tables/table[0-9]+.xml$", xmlFiles, perl = TRUE)]
  tableRelsXML       <- xmlFiles[grepl("table[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  queryTablesXML     <- xmlFiles[grepl("queryTable[0-9]+.xml$", xmlFiles, perl = TRUE)]
  connectionsXML     <- xmlFiles[grepl("connections.xml$", xmlFiles, perl = TRUE)]
  extLinksXML        <- xmlFiles[grepl("externalLink[0-9]+.xml$", xmlFiles, perl = TRUE)]
  extLinksRelsXML    <- xmlFiles[grepl("externalLink[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  
  
  # pivot tables
  pivotTableXML      <- xmlFiles[grepl("pivotTable[0-9]+.xml$", xmlFiles, perl = TRUE)]
  pivotTableRelsXML  <- xmlFiles[grepl("pivotTable[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  pivotDefXML        <- xmlFiles[grepl("pivotCacheDefinition[0-9]+.xml$", xmlFiles, perl = TRUE)]
  pivotDefRelsXML    <- xmlFiles[grepl("pivotCacheDefinition[0-9]+.xml.rels$", xmlFiles, perl = TRUE)]
  pivotRecordsXML    <- xmlFiles[grepl("pivotCacheRecords[0-9]+.xml$", xmlFiles, perl = TRUE)]
  
  ## slicers
  slicerXML          <- xmlFiles[grepl("slicer[0-9]+.xml$", xmlFiles, perl = TRUE)]
  slicerCachesXML    <- xmlFiles[grepl("slicerCache[0-9]+.xml$", xmlFiles, perl = TRUE)]
  
  ## VBA Macro
  vbaProject         <- xmlFiles[grepl("vbaProject\\.bin$", xmlFiles, perl = TRUE)]
  
  ## remove all EXCEPT media and charts
  on.exit(expr = unlink(xmlFiles[!grepl("charts|media|vmlDrawing|comment", xmlFiles, ignore.case = TRUE)], recursive = TRUE, force = TRUE), add = TRUE)
  
  nSheets <- length(worksheetsXML) + length(chartSheetsXML)
  
  ## get Rid of chartsheets, these do not have a worksheet/sheeti.xml
  worksheet_rId_mapping <- NULL
  workbookRelsXML <- xmlFiles[grepl("workbook.xml.rels$", xmlFiles, perl = TRUE)]
  if(length(workbookRelsXML) > 0){
    workbookRelsXML <- paste(readLines(con = workbookRelsXML, encoding="UTF-8", warn = FALSE), collapse = "")
    workbookRelsXML <- .Call("openxlsx_getChildlessNode", workbookRelsXML, "<Relationship ", PACKAGE="openxlsx")
    worksheet_rId_mapping <- workbookRelsXML[grepl("worksheets/sheet", workbookRelsXML, fixed = TRUE)]
  }
  
  ## 
  chartSheetRIds <- NULL
  if(length(chartSheetsXML) > 0){
    
    workbookRelsXML <- workbookRelsXML[grepl("chartsheets/sheet", workbookRelsXML, fixed = TRUE)]
    
    chartSheetRIds <- unlist(getId(workbookRelsXML))
    chartsheet_rId_mapping <- unlist(regmatches(workbookRelsXML, gregexpr('sheet[0-9]+\\.xml', workbookRelsXML, perl = TRUE, ignore.case = TRUE)))
    
    sheetNo <- as.integer(regmatches(chartSheetsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml)", chartSheetsXML, perl = TRUE)))
    chartSheetsXML <- chartSheetsXML[order(sheetNo)]
    
    chartSheetsRelsXML <- xmlFiles[grepl("xl/chartsheets/_rels", xmlFiles, perl = TRUE)]
    sheetNo2 <- as.integer(regmatches(chartSheetsRelsXML, regexpr("(?<=sheet)[0-9]+(?=\\.xml\\.rels)", chartSheetsRelsXML, perl = TRUE)))
    chartSheetsRelsXML <- chartSheetsRelsXML[order(sheetNo2)]
    
    chartSheetsRelsDir <- dirname(chartSheetsRelsXML[1])
    
    
  }
  
  
  ## xl\
  ## xl\workbook
  if(length(workbookXML) > 0){
    
    workbook <- readLines(workbookXML, warn=FALSE, encoding="UTF-8")
    workbook <-  removeHeadTag(workbook)
    
    sheets <- unlist(regmatches(workbook, gregexpr("<sheet .*/sheets>", workbook, perl = TRUE)))
    
    ## sheetId is meaningless
    ## sheet rId links to the workbook.xml.resl which links worksheets/sheet(i).xml file
    ## order they appear here gives order of worksheets in xlsx file
    
    sheetrId <- unlist(getRId(sheets))
    sheetId <- unlist(regmatches(sheets, gregexpr('(?<=sheetId=")[0-9]+', sheets, perl = TRUE)))
    sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
    
    is_chart_sheet <- sheetrId %in% chartSheetRIds
    is_visible <- !grepl("hidden",  unlist(strsplit(sheets, split = "<sheet "))[-1])
    if(length(is_visible) != length(sheetrId))
      is_visible <- rep(TRUE, length(sheetrId))
    
    ## add worksheets to wb
    j <- 1
    for(i in 1:length(sheetrId)){
      
      if(is_chart_sheet[i]){
        
        count <- 0
        txt <- paste(readLines(chartSheetsXML[j], warn = FALSE, encoding = "UTF-8"), collapse = "")
        
        zoom <- regmatches(txt, regexpr('(?<=zoomScale=")[0-9]+', txt, perl = TRUE))
        if(length(zoom) == 0)
          zoom <- 100
        
        tabColour <- .Call("openxlsx_getChildlessNode", txt, "<tabColor ", PACKAGE = "openxlsx")
        if(length(tabColour) == 0)
          tabColour <- NULL
        
        j <- j + 1L
        
        wb$addChartSheet(sheetName = sheetNames[i], tabColour = tabColour, zoom = as.numeric(zoom))
      }else{
        wb$addWorksheet(sheetNames[i], visible = is_visible[i])
      }
      
    }
    
    
    ## replace sheetId
    for(i in 1:nSheets)
      wb$workbook$sheets[[i]] <- gsub(sprintf(' sheetId="%s"', i), sprintf(' sheetId="%s"', sheetId[i]), wb$workbook$sheets[[i]])
    
    
    ## additional workbook attributes
    calcPr <- .Call("openxlsx_getChildlessNode", workbook, "<calcPr ", PACKAGE = "openxlsx")
    if(length(calcPr) > 0)
      wb$workbook$calcPr <- calcPr
    
    
    workbookPr <- .Call("openxlsx_getChildlessNode", workbook, "<workbookPr ", PACKAGE = "openxlsx")
    if(length(calcPr) > 0)
      wb$workbook$workbookPr <- workbookPr
    
    
    ## defined Names
    dNames <- .Call("openxlsx_getNodes", workbook, "<definedNames>", PACKAGE = "openxlsx")
    if(length(dNames) > 0){
      dNames <- gsub("^<definedNames>|</definedNames>$", "", dNames)
      wb$workbook$definedNames <- paste0(.Call("openxlsx_getNodes", dNames, "<definedName", PACKAGE = "openxlsx"), ">")
    }
    
    
  }
  
  
  
  
  
  ## xl\sharedStrings
  if(length(sharedStringsXML) > 0){
    
    sharedStrings <- readLines(sharedStringsXML, warn = FALSE, encoding = "UTF-8")
    sharedStrings <- paste(sharedStrings, collapse = "\n")
    sharedStrings <- removeHeadTag(sharedStrings)
    
    uniqueCount <- as.integer(regmatches(sharedStrings, regexpr('(?<=uniqueCount=")[0-9]+', sharedStrings, perl = TRUE)))
    
    ## read in and get <si> nodes
    vals <- .Call("openxlsx_getNodes", sharedStrings, "<si>", PACKAGE = "openxlsx")
    Encoding(vals) <- "UTF-8"
    attr(vals, "uniqueCount") <- uniqueCount
    
    wb$sharedStrings <- vals
    
  }
  
  
  ## xl\pivotTables & xl\pivotCache
  if(length(pivotTableXML) > 0){
    
    # pivotTable cacheId links to workbook.xml which links to workbook.xml.rels via rId
    # we don't modify the cacheId, only the rId
    
    nPivotTables      <- length(pivotTableXML)
    rIds <- 20000L + 1:nPivotTables
    
    pivotTableXML     <- pivotTableXML[order(nchar(pivotTableXML), pivotTableXML)]
    pivotTableRelsXML <- pivotTableRelsXML[order(nchar(pivotTableRelsXML), pivotTableRelsXML)]
    pivotDefXML       <- pivotDefXML[order(nchar(pivotDefXML), pivotDefXML)]
    pivotDefRelsXML   <- pivotDefRelsXML[order(nchar(pivotDefRelsXML), pivotDefRelsXML)]
    pivotRecordsXML   <- pivotRecordsXML[order(nchar(pivotRecordsXML), pivotRecordsXML)]
    
    wb$pivotTables <- character(nPivotTables)
    wb$pivotTables.xml.rels <- character(nPivotTables)
    wb$pivotDefinitions <- character(nPivotTables)
    wb$pivotDefinitionsRels <- character(nPivotTables)
    wb$pivotRecords <- character(nPivotTables)
    
    wb$pivotTables[1:length(pivotTableXML)] <-
      unlist(lapply(pivotTableXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"))))
    
    
    wb$pivotTables.xml.rels[1:length(pivotTableRelsXML)] <-
      unlist(lapply(pivotTableRelsXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"))))
    
    wb$pivotDefinitions[1:length(pivotDefXML)]  <-
      unlist(lapply(pivotDefXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"))))
    
    
    wb$pivotDefinitionsRels[1:length(pivotDefRelsXML)] <-
      unlist(lapply(pivotDefRelsXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"))))
    
    
    wb$pivotRecords[1:length(pivotRecordsXML)] <-
      unlist(lapply(pivotRecordsXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"))))
    
    ## update content_types
    wb$Content_Types <- c(wb$Content_Types, unlist(lapply(1:nPivotTables, contentTypePivotXML)))
    
    ## workbook rels
    wb$workbook.xml.rels <- c(wb$workbook.xml.rels,    
                              sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition%s.xml"/>', rIds, 1:nPivotTables)
    )
    
    caches <- .Call("openxlsx_getChildlessNode", workbook, "<pivotCache ", PACKAGE = "openxlsx")
    for(i in 1:length(caches))
      caches[i] <- gsub('"rId[0-9]+"', sprintf('"rId%s"', rIds[i]), caches[i])
    
    wb$workbook$pivotCaches <- paste0('<pivotCaches>', paste(caches, collapse = ""), '</pivotCaches>')
    
  }
  
  ## xl\vbaProject
  if(length(vbaProject) > 0){
    wb$vbaProject <- vbaProject
    wb$Content_Types[grepl('<Override PartName="/xl/workbook.xml" ', wb$Content_Types)] <- '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>'
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>')    
  }
  
  
  ## xl\styles
  if(length(stylesXML) > 0){
    styleObjects <- wb$loadStyles(stylesXML)
  }else{
    styleObjects <- list()
  }
  
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
    
    chartNames <- basename(charts)
    nCharts <- sum(grepl("chart[0-9]+.xml", chartNames))
    nChartStyles <- sum(grepl("style[0-9]+.xml", chartNames))
    nChartCol <- sum(grepl("colors[0-9]+.xml", chartNames))
    
    if(nCharts > 0)
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/chart%s.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', 1:nCharts))
    
    if(nChartStyles > 0)
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/style%s.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>', 1:nChartStyles))
    
    if(nChartCol > 0)
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/charts/colors%s.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>', 1:nChartCol))
    
    if(length(chartsRels)){
      charts <- c(charts, chartsRels)
      chartNames <- c(chartNames, file.path("_rels", basename(chartsRels)))
    }
    
    names(charts) <- chartNames
    wb$charts <- charts
    
  }
  
  
  
  
  
  
  ## xl\theme
  if(length(themeXML) > 0)
    wb$theme <- removeHeadTag(paste(unlist(lapply(sort(themeXML)[[1]], function(x) readLines(x, warn = FALSE, encoding = "UTF-8"))), collapse = ""))
  
  
  ## externalLinks
  if(length(extLinksXML) > 0){
    wb$externalLinks <- lapply(sort(extLinksXML), function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
    
    wb$Content_Types <-c(wb$Content_Types, 
                         sprintf('<Override PartName="/xl/externalLinks/externalLink%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>', 1:length(extLinksXML)))
    
    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf('<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>',
                                                            1:length(extLinksXML)))
  }
  
  ## externalLinksRels
  if(length(extLinksRelsXML) > 0)
    wb$externalLinksRels <- lapply(sort(extLinksRelsXML), function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
  
  
  
  
  
  
  
  ##*----------------------------------------------------------------------------------------------*##
  ### BEGIN READING IN WORKSHEET DATA
  ##*----------------------------------------------------------------------------------------------*##
  
  ## xl\worksheets
  file_names <- regmatches(worksheet_rId_mapping, regexpr("sheet[0-9]+\\.xml", worksheet_rId_mapping, perl = TRUE))
  file_rIds <- unlist(getId(worksheet_rId_mapping))
  file_names <- file_names[match(sheetrId, file_rIds)]
  
  worksheetsXML <- file.path(dirname(worksheetsXML), file_names)
  wb <- .Call("openxlsx_loadworksheets", wb, styleObjects, worksheetsXML, is_chart_sheet)
  
  ## Fix styleobject encoding
  if(length(wb$styleObjects) > 0){
    style_names <- sapply(wb$styleObjects, "[[", "sheet")
    Encoding(style_names) <- "UTF-8"
    wb$styleObjects <- lapply(1:length(style_names), function(i) {wb$styleObjects[[i]]$sheet = style_names[[i]]; wb$styleObjects[[i]]})
  }
  
  ##*----------------------------------------------------------------------------------------------*##
  ### READING IN WORKSHEET DATA COMPLETE
  ##*----------------------------------------------------------------------------------------------*##
  
  
  
  
  
  ## Next sheetRels to see which drawings_rels belongs to which sheet
  
  
  if(length(sheetRelsXML) > 0){
    
    ## sheetrId is order sheet appears in xlsx file
    ## create a 1-1 vector of rels to worksheet
    ## have rels is boolean vector where i-the element is TRUE/FALSE if sheet has a rels sheet
    
    if(length(chartSheetsXML) == 0){
      allRels <- file.path(dirname(sheetRelsXML[1]), paste0(file_names, ".rels"))
      haveRels <- allRels %in% sheetRelsXML
    }else{
      
      haveRels <- rep(FALSE, length(wb$worksheets))
      allRels <- rep("", length(wb$worksheets))
      
      for(i in 1:nSheets){
        if(is_chart_sheet[i]){
          ind <- which(chartSheetRIds == sheetrId[i])
          rels_file <- file.path(chartSheetsRelsDir, paste0(chartsheet_rId_mapping[ind], ".rels"))
        }else{
          ind <- sheetrId[i]
          rels_file <- file.path(xmlDir, "xl", "worksheets", "_rels", paste0(file_names[i], ".rels"))
        }
        if(file.exists(rels_file)){
          allRels[i] <- rels_file
          haveRels[i] <- TRUE
        }
      }
    }
    
    ## sheet.xml have been reordered to be in the order of sheetrId
    ## not every sheet has a worksheet rels
    
    xml <- lapply(1:length(allRels), function(i) {
      if(haveRels[i]){
        
        xml <- readLines(allRels[[i]], warn = FALSE, encoding = "UTF-8")
        xml <- removeHeadTag(xml)
        xml <- gsub("<Relationships .*?>", "", xml)
        xml <- gsub("</Relationships>", "", xml)
        xml <- .Call("openxlsx_getChildlessNode", xml, "<Relationship ", PACKAGE ="openxlsx")
        
      }else{
        xml <- "<Relationship >"
      }
      return(xml)
    })
    
    
    if(length(slicerXML) > 0){
      
      slicerXML <- slicerXML[order(nchar(slicerXML), slicerXML)]
      slicersFiles <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=slicer)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      inds <- sapply(slicersFiles, length) > 0
      
      
      ## worksheet_rels Id for slicer will be rId0
      k <- 1L
      wb$slicers <- rep("", nSheets)
      for(i in 1:nSheets){
        
        ## read in slicer[j].XML sheets into sheet[i]
        if(inds[i]){
          
          wb$slicers[[i]] <- removeHeadTag(.Call("openxlsx_cppReadFile", slicerXML[k], PACKAGE = "openxlsx"))
          k <- k + 1L
          
          wb$worksheets_rels[[i]] <- unlist(c(wb$worksheets_rels[[i]],
                                              sprintf('<Relationship Id="rId0" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer%s.xml"/>', i)))
          wb$Content_Types <- c(wb$Content_Types,
                                sprintf('<Override PartName="/xl/slicers/slicer%s.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', i))
          
          ## Append slicer to worksheet extLst
          wb$worksheets[[i]]$extLst <- c(wb$worksheets[[i]]$extLst, genBaseSlicerXML())
          
        }
      }
      
    }
    
    
    if(length(slicerCachesXML) > 0){
      
      ## ---- slicerCaches
      inds <- 1:length(slicerCachesXML)
      wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/slicerCaches/slicerCache%s.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>', inds))
      wb$slicerCaches <- sapply(slicerCachesXML[order(nchar(slicerCachesXML), slicerCachesXML)], function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
      wb$workbook.xml.rels <- c(wb$workbook.xml.rels, sprintf('<Relationship Id="rId%s" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache%s.xml"/>', 1E5 + inds, inds))
      wb$workbook$extLst <- c(wb$workbook$extLst,  genSlicerCachesExtLst(1E5 + inds))
      
      
    }    
    
    ## tables
    if(length(tablesXML) > 0){
      
      tables <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=table)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      tableSheets <- unlist(lapply(1:length(sheetrId), function(i) rep(i, length(tables[[i]]))))  
      
      if(length(unlist(tables)) > 0){
        ## get the tables that belong to each worksheet and create a worksheets_rels for each
        tCount <- 2L ## table r:Ids start at 3
        for(i in 1:length(tables)){
          if(length(tables[[i]]) > 0){
            k <- 1:length(tables[[i]]) + tCount
            wb$worksheets_rels[[i]] <- unlist(c(wb$worksheets_rels[[i]],
                                                sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table%s.xml"/>',  k,  k)))
            
            
            wb$worksheets[[i]]$tableParts <- sprintf("<tablePart r:id=\"rId%s\"/>", k)
            tCount <- tCount + length(k)
          }
        }
        
        ## sort the tables into the order they appear in the xml and tables variables
        names(tablesXML) <- basename(tablesXML)
        tablesXML <- tablesXML[sprintf("table%s.xml", unlist(tables))]
        
        ## tables are now in correct order so we can read them in as they are
        wb$tables <- sapply(tablesXML, function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx")))
        
        ## pull out refs and attach names
        refs <- regmatches(wb$tables, regexpr('(?<=ref=")[0-9A-Z:]+', wb$tables, perl = TRUE))
        names(wb$tables) <- refs
        
        wb$Content_Types <- c(wb$Content_Types, sprintf('<Override PartName="/xl/tables/table%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>', 1:length(wb$tables)+2))   
        
        ## relabel ids
        for(i in 1:length(wb$tables)){
          newId <- sprintf(' id="%s" ', i+2)
          wb$tables[[i]] <- sub(' id="[0-9]+" ' , newId, wb$tables[[i]])
        }
        
        displayNames <- unlist(regmatches(wb$tables, regexpr('(?<=displayName=").*?[^"]+', wb$tables, perl = TRUE)))
        if(length(displayNames) != length(tablesXML))
          displayNames <- paste0("Table", 1:length(tablesXML))
        
        attr(wb$tables, "sheet") <- tableSheets
        attr(wb$tables, "tableName") <- displayNames
        
      }
    } ## if(length(tablesXML) > 0)
    
    ## hyperlinks
    hlinks <- lapply(xml, function(x) x[grepl("hyperlink", x) & grepl("External", x)])
    hlinksInds <- which(sapply(hlinks, length) > 0)
    
    if(length(hlinksInds) > 0){
      
      hlinks <- hlinks[hlinksInds]
      for(i in 1:length(hlinksInds)){
        targets <- unlist(lapply(hlinks[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
        targets <- gsub('"$', "", targets)
        
        names(wb$hyperlinks[[hlinksInds[i]]]) <- targets  
      }  
    }
    
    ## xml is in the order of the sheets, drawIngs is toes to sheet position of hasDrawing
    ## Not every sheet has a drawing.xml
    
    ## drawings
    
    drawXMLrelationship <- lapply(xml, function(x) x[grepl("drawings/drawing", x)])
    hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing
    
    if(length(drawingRelsXML) > 0){
      
      dRels <- lapply(drawingRelsXML, readLines, warn = FALSE)
      dRels <- unlist(lapply(dRels, removeHeadTag))
      dRels <- gsub("<Relationships .*?>", "", dRels)
      dRels <- gsub("</Relationships>", "", dRels)
    }
    
    if(length(drawingsXML) > 0){
      dXML <- lapply(drawingsXML, readLines, warn = FALSE, encoding = "UTF-8")  
      dXML <- unlist(lapply(dXML, removeHeadTag))
      dXML <- gsub("<xdr:wsDr .*?>", "", dXML)
      dXML <- gsub("</xdr:wsDr>", "", dXML)
      
      ## split at one/two cell Anchor
      dXML <- regmatches(dXML, gregexpr("<xdr:(oneCell|twoCell|absolute)Anchor.*?</xdr:(oneCell|twoCell|absolute)Anchor>", dXML))
    }
    
    
    ## loop over all worksheets and assign drawing to sheet
    if(any(hasDrawing)){
      for(i in 1:length(xml)){
        
        if(hasDrawing[i]){
          
          target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))
          
          ## sheet_i has which(hasDrawing)[[i]]
          relsInd <- grepl(target, drawingRelsXML)
          if(any(relsInd))
            wb$drawings_rels[i] <- dRels[relsInd]
          
          drawingInd <- grepl(target, drawingsXML)
          if(any(drawingInd))
            wb$drawings[i] <- dXML[drawingInd]
          
        }
      }
    }
    
    
    
    ## vmlDrawing and comments
    if(length(commentsXML) > 0){
      
      wb$Content_Types <- c(wb$Content_Types, '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>')
      
      drawXMLrelationship <- lapply(xml, function(x) x[grepl("drawings/vmlDrawing[0-9]+\\.vml", x)])
      hasDrawing <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing
      
      commentXMLrelationship <- lapply(xml, function(x) x[grepl("comments[0-9]+\\.xml", x)])
      hasComment <- sapply(drawXMLrelationship, length) > 0 ## which sheets have a drawing
      
      for(i in 1:length(xml)){
        
        if(hasComment[i]){
          
          target <- unlist(lapply(drawXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))
          
          txt <- paste(readLines(vmlDrawingXML[grepl(target, vmlDrawingXML)], warn = FALSE), collapse = "\n")
          txt <- removeHeadTag(txt)
          
          cd <- .Call("openxlsx_getNodes", txt, "<x:ClientData", PACKAGE = "openxlsx")
          cd <- cd[grepl('ObjectType="Note"', cd)]
          cd <- paste0(cd, ">")

          
          target <- unlist(lapply(commentXMLrelationship[[i]], function(x) regmatches(x, gregexpr('(?<=Target=").*?"', x, perl = TRUE))[[1]]))
          target <- basename(gsub('"$', "", target))
          
          txt <- paste(readLines(commentsXML[grepl(target, commentsXML)], warn = FALSE), collapse = "\n")
          txt <- removeHeadTag(txt)
          
          authors <- .Call("openxlsx_getNodes", txt, "<author>", PACKAGE = "openxlsx")
          authors <- gsub("<author>|</author>", "", authors)
          
          comments <- .Call("openxlsx_getNodes", txt, "<commentList>", PACKAGE = "openxlsx")
          comments <- gsub( "<commentList>", "", comments)
          comments <- .Call("openxlsx_getNodes", comments, "<comment", PACKAGE = "openxlsx")
          
          refs <- regmatches(comments, regexpr('(?<=ref=").*?[^"]+', comments, perl = TRUE))
          
          authorsInds <- as.integer(regmatches(comments, regexpr('(?<=authorId=").*?[^"]+', comments, perl = TRUE))) + 1
          authors <- authors[authorsInds]
          
          style <- lapply(comments, function(txt) .Call("openxlsx_getNodes", txt, "<rPr>", PACKAGE = "openxlsx"))
          
          comments <- regmatches(comments, gregexpr('(?<=<t( |>)).*?[^/]+', comments, perl = TRUE))
          comments <- lapply(comments, function(x) gsub("<", "", x))
          comments <- lapply(comments, function(x) gsub(".*?>", "", x, perl = TRUE))
          
          
          wb$comments[[i]] <- lapply(1:length(comments), function(j){
            
            comment_list <- list("ref" = refs[j],
                                 "author" = authors[j],
                                 "comment" = comments[[j]],
                                 "style"  = style[[j]],
                                 "clientData" = cd[[j]])    
            
          })
          
          
          
        }
      }
    }
    
    
    
    
    ## pivot tables
    if(length(pivotTableXML) > 0){
      
      pivotTableJ <- lapply(xml, function(x) as.integer(regmatches(x, regexpr("(?<=pivotTable)[0-9]+(?=\\.xml)", x, perl = TRUE))))
      sheetWithPivot <- which(sapply(pivotTableJ, length) > 0)
      
      pivotRels <- lapply(xml, function(x) {y <- x[grepl("pivotTable", x)]; y[order(nchar(y), y)]})
      hasPivot <- sapply(pivotRels, length) > 0
      
      ## Modify rIds
      for(i in 1:length(pivotRels)){
        if(hasPivot[i]){
          for(j in 1:length(pivotRels[[i]]))  
            pivotRels[[i]][j] <- gsub('"rId[0-9]+"', sprintf('"rId%s"', 20000L + j), pivotRels[[i]][j])
          
          wb$worksheets_rels[[i]] <- c(wb$worksheets_rels[[i]] , pivotRels[[i]])
        }
      }  
      
      
      ## remove any workbook_res references to pivot tables that are not being used in worksheet_rels
      inds <- 1:length(wb$pivotTables.xml.rels)
      fileNo <- as.integer(unlist(regmatches(unlist(wb$worksheets_rels), gregexpr('(?<=pivotTable)[0-9]+(?=\\.xml)', unlist(wb$worksheets_rels), perl = TRUE))))
      inds <- inds[!inds %in% fileNo]
      
      if(length(inds) > 0){
        
        toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", inds), collapse = "|")    
        fileNo <- which(grepl(toRemove, wb$pivotTables.xml.rels))
        toRemove <- paste(sprintf("(pivotCacheDefinition%s\\.xml)", fileNo), collapse = "|")
        
        ## remove reference to file from workbook.xml.res
        wb$workbook.xml.rels <- wb$workbook.xml.rels[!grepl(toRemove, wb$workbook.xml.rels)]
      }
    }
    
  } ## end of worksheetRels
  
  
  
  ## queryTables
  if(length(queryTablesXML) > 0){
    
    ids <- as.numeric(regmatches(queryTablesXML, regexpr("[0-9]+(?=\\.xml)", queryTablesXML, perl = TRUE)))
    wb$queryTables <- unlist(lapply(queryTablesXML[order(ids)], function(x) removeHeadTag(.Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"))))
    wb$Content_Types <- c(wb$Content_Types, 
                          sprintf('<Override PartName="/xl/queryTables/queryTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>', 1:length(queryTablesXML)))   
  }
  
  
  ## connections
  if(length(connectionsXML) > 0){
    wb$connections <- removeHeadTag(.Call("openxlsx_cppReadFile", connectionsXML, PACKAGE = "openxlsx"))
    wb$workbook.xml.rels <- c(wb$workbook.xml.rels, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>')
    wb$Content_Types <- c(wb$Content_Types, '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>')
  }
  
  
  
  
  ## table rels
  if(length(tableRelsXML) > 0){
    
    ## table_i_might have tableRels_i but I am re-ordering the tables to be in order of worksheets
    ## I make every table have a table_rels so i need to fill in the gaps if any table_rels are missing
    
    tmp <- paste0(basename(tablesXML), ".rels")
    hasRels <- tmp %in% basename(tableRelsXML)
    
    ## order tableRelsXML
    tableRelsXML <- tableRelsXML[match(tmp[hasRels], basename(tableRelsXML))]
    
    ##
    wb$tables.xml.rels <- character(length=length(tablesXML))
    
    ## which sheet does it belong to
    xml <- sapply(tableRelsXML, function(x) .Call("openxlsx_cppReadFile", x, PACKAGE = "openxlsx"), USE.NAMES = FALSE)
    xml <- sapply(xml, removeHeadTag, USE.NAMES = FALSE)
    
    wb$tables.xml.rels[hasRels] <- xml
    
  }else if(length(tablesXML) > 0){
    wb$tables.xml.rels <- rep("", length(tablesXML))
  }
  
  
  
  
  
  return(wb)
  
}


