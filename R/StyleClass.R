


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
  numFmt <<- NULL
  fill <<- NULL
  wrapText <<- NULL
})



mergeStyle = function(oldStyle, newStyle){

  ## This function is used to merge an existing cell style with a new style to create a stacked style.
  
  oldStyle <- oldStyle$copy()
  
  if(!is.null(newStyle$fontName))
    oldStyle$fontName <- newStyle$fontName
  
  if(!is.null(newStyle$fontColour))
    oldStyle$fontColour <- newStyle$fontColour

  if(!is.null(newStyle$fontSize))
    oldStyle$fontSize <- newStyle$fontSize
  
  if(!is.null(newStyle$fontFamily))
    oldStyle$fontFamily <- newStyle$fontFamily
  
  if(!is.null(newStyle$fontScheme))
    oldStyle$fontScheme <- newStyle$fontScheme
  
  if(!is.null(newStyle$fontDecoration))
    oldStyle$fontDecoration <- newStyle$fontDecoration
  
  ## borders
  if(!is.null(newStyle$borderTop))
    oldStyle$borderTop <- newStyle$borderTop
  
  if(!is.null(newStyle$borderLeft))
    oldStyle$borderLeft <- newStyle$borderLeft
  
  if(!is.null(newStyle$borderRight))
    oldStyle$borderRight <- newStyle$borderRight
  
  if(!is.null(newStyle$borderBottom))
    oldStyle$borderBottom <- newStyle$borderBottom
  
  if(!is.null(newStyle$borderTopColour))
    oldStyle$borderTopColour <- newStyle$borderTopColour
  
  if(!is.null(newStyle$borderLeftColour))
    oldStyle$borderLeftColour <- newStyle$borderLeftColour
  
  if(!is.null(newStyle$borderRightColour))
    oldStyle$borderRightColour <- newStyle$borderRightColour
  
  if(!is.null(newStyle$borderBottomColour))
    oldStyle$borderBottomColour <- newStyle$borderBottomColour
  
  ## other
  if(!is.null(newStyle$halign))
    oldStyle$halign <- newStyle$halign
  
  if(!is.null(newStyle$valign))
    oldStyle$valign <- newStyle$valign
  
  if(!is.null(newStyle$textRotation))
    oldStyle$textRotation <- newStyle$textRotation
  
  if(!is.null(newStyle$numFmt))
    oldStyle$numFmt <- newStyle$numFmt
  
  if(!is.null(newStyle$fill))
    oldStyle$fill <- newStyle$fill
  
  if(!is.null(newStyle$wrapText))
    oldStyle$wrapText <- newStyle$wrapText
    
    
    return(oldStyle)

}





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
  
  if(!is.null(numFmt)){
    if(as.integer(numFmt$numFmtId) %in% unlist(numFmtMapping)){
      numFmtStr <- validNumFmt[unlist(numFmtMapping) == as.integer(numFmt$numFmtId)]
    }else{
      numFmtStr <- sprintf('"%s"', numFmt$formatCode)
    }
  }else{
    numFmtStr <- "GENERAL"
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
