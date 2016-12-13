



Workbook <- setRefClass("Workbook", 
                        
                        fields = c("sheet_names" = "character",
                                   
                                   "charts" = "ANY",
                                   "isChartSheet" = "logical",
                                   
                                   "colWidths" = "ANY",
                                   "connections" = "ANY",
                                   "Content_Types" = "character",
                                   "core" = "character",
                                   "drawings" = "ANY",
                                   "drawings_rels" = "ANY",
                                   "embeddings" = "ANY",
                                   "externalLinks" = "ANY",
                                   "externalLinksRels" = "ANY",
                                   
                                   "headFoot" = "ANY",
                                   "media" = "ANY",
                                   
                                   "pivotTables" = "ANY",
                                   "pivotTables.xml.rels" = "ANY",
                                   "pivotDefinitions" = "ANY",
                                   "pivotRecords" = "ANY",
                                   "pivotDefinitionsRels" = "ANY",
                                   
                                   "queryTables" = "ANY",
                                   "rowHeights" = "ANY",
                                   
                                   "slicers" = "ANY",
                                   "slicerCaches" = "ANY",
                                   
                                   "sharedStrings" = "ANY",
                                   "styleObjects" = "ANY",
                                   
                                   "styles" = "ANY",
                                   "styleInds" = "ANY",
                                   "tables" = "ANY",
                                   "tables.xml.rels" = "ANY",
                                   "theme" = "ANY",
                                   
                                   "vbaProject" = "ANY",
                                   "vml" = "ANY",
                                   "vml_rels" = "ANY",
                                   "comments" = "ANY",
                                   
                                   "workbook" = "ANY",
                                   "workbook.xml.rels" = "ANY",
                                   "worksheets" = "ANY",
                                   "worksheets_rels" = "ANY",
                                   "sheetOrder" = "integer",
                                   "dataCount" = "numeric")
)



Style <- setRefClass("Style", 
                     
                     fields = c("fontName",
                                "fontColour",
                                "fontSize",
                                "fontFamily",
                                "fontScheme",
                                "fontDecoration",
                                "borderTop",
                                "borderLeft",
                                "borderRight",
                                "borderBottom",
                                "borderTopColour",
                                "borderLeftColour",
                                "borderRightColour",
                                "borderBottomColour",
                                "halign",
                                "valign",
                                "indent",
                                "textRotation",
                                "numFmt",
                                "fill",
                                "wrapText",
                                "xfId"),
                     methods = list()
)

# SheetData <- setRefClass("SheetData", 
#                          
#                          fields = c("row" = "integer",
#                                     "col" = "integer",
#                                     "t" = "integer",
#                                     "v" = "numeric",
#                                     "f" = "character",
#                                     "dataCount" = "integer",
#                                     "n_elements" = "integer")
#                          
# )
# 




WorkSheet <- setRefClass("WorkSheet", 
                         
                         fields = c("sheetPr" = "character",
                                    "dimension" = "character",
                                    "sheetViews" =  "character",
                                    "sheetFormatPr" = "character",
                                    "cols" = "character",
                                    
                                    "sheetData" = "ANY",
                                    
                                    "autoFilter" = "character",
                                    "mergeCells" = "ANY",
                                    "conditionalFormatting" = "character",
                                    "dataValidations" = "ANY", 
                                    
                                    "freezePane" = "character",
                                    "hyperlinks" = "ANY",
                                    
                                    "pageMargins" = "character",
                                    "pageSetup" = "character",
                                    "headerFooter" = "ANY",
                                    "rowBreaks" = "character",
                                    "colBreaks" = "character",
                                    "drawing" = "character",
                                    "legacyDrawing" = "character",
                                    "legacyDrawingHF" = "character",
                                    "oleObjects" = "character",
                                    "tableParts" = "character",
                                    "extLst" = "character")
)




ChartSheet <- setRefClass("ChartSheet", 
                         
                         fields = c("sheetPr" = "character",
                                    "sheetViews" =  "character",
                                    "pageMargins" = "character",
                                    "drawing" = "character",
                                    "hyperlinks" = "ANY")
)


