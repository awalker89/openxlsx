


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
                                "textRotation",
                                "numFmt",
                                "fill",
                                "wrapText"),
  methods = list()
)


Style$methods(initialize = function(){
  
  fontName <<- NULL
  fontColour <<- NULL
  fontSize <<- NULL
  fontFamily <<- NULL
  fontScheme <<- NULL
  fontDecoration <<- NULL
  
  borderTop <<- NULL
  borderLeft <<- NULL
  borderRight <<- NULL
  borderBottom <<- NULL
  borderTopColour <<- NULL
  borderLeftColour <<- NULL
  borderRightColour <<- NULL
  borderBottomColour <<- NULL
  
  halign <<- NULL
  valign <<- NULL
  textRotation <<- NULL
  numFmt <<- list("numFmtId" = 0)
  fill <<- NULL
  wrapText <<- NULL
})


Style$methods(show = function(print = TRUE){
 
  numFmtMapping <- list(list("numFmtId" = 0),
                        list("numFmtId" = 2),
                        list("numFmtId" = 164),
                        list("numFmtId" = 44),
                        list("numFmtId" = 14),
                        list("numFmtId" = 167),
                        list("numFmtId" = 10),
                        list("numFmtId" = 11),
                        list("numFmtId" = 49))
  
  validNumFmt <- c("GENERAL", "NUMBER", "CURRENCY", "ACCOUNTING", "DATE", "TIME", "PERCENTAGE", "SCIENTIFIC", "TEXT")
  
  if(as.integer(numFmt$numFmtId) %in% unlist(numFmtMapping)){
    numFmtStr <- validNumFmt[unlist(numFmtMapping) == as.integer(numFmt$numFmtId)]
  }else{
    numFmtStr <- sprintf('"%s"', numFmt$formatCode)
  }
  
  borders <- c(sprintf("Top: %s", borderTop), sprintf("Bottom: %s", borderBottom), sprintf("Left: %s", borderLeft), sprintf("Right: %s", borderRight))
  borderColours <- gsub("^FF", "#", c(borderTopColour, borderBottomColour, borderLeftColour, borderRightColour))

  fgFill <- fill$fillFg
  bgFill <- fill$fillBg

  styleShow <- "A custom cell style. \n\n"

  styleShow <- append(styleShow, sprintf("Cell formatting: %s \n", numFmtStr))  ## numFmt
  styleShow <- append(styleShow, sprintf("Font name: %s \n", fontName[[1]]))  ## Font name
  styleShow <- append(styleShow, sprintf("Font size: %s \n", fontSize[[1]]))  ## Font size
  styleShow <- append(styleShow, sprintf("Font colour: %s \n", gsub("^FF", "#", fontColour[[1]])))  ## Font colour
  
  ## Font decoration
  if(length(fontDecoration) > 0)
    styleShow <- append(styleShow, sprintf("Font decoration: %s \n", paste(fontDecoration, collapse = ", ")))
  
  if(length(borders) > 0){
    styleShow <- append(styleShow, sprintf("Cell borders: %s \n", paste(borders, collapse = ", ")))  ## Cell borders
    styleShow <- append(styleShow, sprintf("Cell border colours: %s \n", paste(borderColours, collapse = ", ")))  ## Cell borders
  }
  
  if(!is.null(halign))
    styleShow <- append(styleShow, sprintf("Cell horz. align: %s \n", halign))  ## Cell horizontal alignment
  
  if(!is.null(halign))
    styleShow <- append(styleShow, sprintf("Cell vert. align: %s \n", valign))  ## Cell vertical alignment 
  
  if(!is.null(textRotation))
    styleShow <- append(styleShow, sprintf("Cell text rotation: %s \n", textRotation))  ## Cell text rotation 
  
  ## Cell fill colour
  if(length(fgFill) > 0)
    styleShow <- append(styleShow, sprintf("Cell fill foreground: %s \n", paste(paste0(names(fgFill),": ", sub("^FF", "#", fgFill)), collapse = ", ")))
  
  if(length(bgFill) > 0)
    styleShow <- append(styleShow, sprintf("Cell fill background: %s \n", paste(paste0(names(bgFill),": ", sub("^FF", "#", bgFill)), collapse = ", ")))
  
  styleShow <- append(styleShow, sprintf("wraptext: %s", wrapText))  ## wrap text
  
  styleShow <- c(styleShow, "\n\n")

  if(print)
    cat(styleShow)
  
  return(invisible(styleShow))
  
})
  

          

Style$methods(as.list = function(){
  
  l <- list(
  "fontName" = fontName,
  "fontColour" = fontColour,
  "fontSize" = fontSize,
  "fontFamily" = fontFamily,
  "fontScheme" = fontScheme,
  "fontDecoration" = fontDecoration,
  
  "borderTop" = borderTop,
  "borderLeft" = borderLeft,
  "borderRight" = borderRight,
  "borderBottom" = borderBottom,
  "borderTopColour" = borderTopColour,
  "borderLeftColour" = borderLeftColour,
  "borderRightColour" = borderRightColour,
  "borderBottomColour" = borderBottomColour,
  
  "halign" = halign,
  "valign" = valign,
  "textRotation" = textRotation,
  "numFmt" = numFmt,
  "fillFg" = fill$fillFg,
  "fillBg" = fill$fillBg,
  "wrapText" = wrapText
  )
  
  l[sapply(l, length) > 0]
  
})
