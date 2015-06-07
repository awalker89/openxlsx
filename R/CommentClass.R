


Comment <- setRefClass("Comment", 
                       
                       fields = c("text",
                                  "author",
                                  "style",
                                  "width",
                                  "height"),
                       
                       methods = list()
)


Comment$methods(initialize = function(text, author, style, width = 2, height = 4){
  
  text <<- text
  author <<- author
  style <<- style
  width <<- width
  height <<- height
  
})


Comment$methods(show = function(){
  
  showText <- sprintf("Author: %s\n", author)
  showText <- c(showText, sprintf("Text:\n %s\n\n", paste(text, collapse = "")))
  styleShow <- "Style:\n"
  
  if("list" %in% class(style)){

    for(i in 1:length(style)){
      
      styleShow <- append(styleShow, sprintf("Font name: %s\n", style[[i]]$fontName[[1]]))  ## Font name
      styleShow <- append(styleShow, sprintf("Font size: %s\n", style[[i]]$fontSize[[1]]))  ## Font size
      styleShow <- append(styleShow, sprintf("Font colour: %s\n", gsub("^FF", "#",  style[[i]]$fontColour[[1]])))  ## Font colour
      
      ## Font decoration
      if(length(style[[i]]$fontDecoration) > 0)
        styleShow <- append(styleShow, sprintf("Font decoration: %s\n", paste(style[[i]]$fontDecoration, collapse = ", ")))
      
      styleShow <- append(styleShow, "\n\n")
    }
    
  }else{
    
    styleShow <- append(styleShow, sprintf("Font name: %s \n", style$fontName[[1]]))  ## Font name
    styleShow <- append(styleShow, sprintf("Font size: %s \n", style$fontSize[[1]]))  ## Font size
    styleShow <- append(styleShow, sprintf("Font colour: %s \n", gsub("^FF", "#",  style$fontColour[[1]])))  ## Font colour
    
    ## Font decoration
    if(length(style$fontDecoration) > 0)
      styleShow <- append(styleShow, sprintf("Font decoration: %s \n", paste(style$fontDecoration, collapse = ", ")))
    
    styleShow <- append(styleShow, "\n\n")
    
  }
  
  showText <- paste0(paste(showText, collapse = ""), paste(styleShow, collapse = ""), collapse = "")
  cat(showText)

})



#' @name createComment
#' @title write a cell comment
#' @param comment Comment text. Character vector of length 1
#' @param author Author of comment. Character vector of length 1
#' @param style A Style object. See \code{\link{createStyle}}.
#' @param width Textbox integer width in number of cells
#' @param height Textbox integer height in number of cells
#' @export
#' @seealso \code{\link{writeComment}}
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' 
#' c1 <- createComment(comment = "this is comment")
#' writeComment(wb, 1, col = "B", row = 10, comment = c1)
#' 
#' s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = c("BOLD"))
#' s2 <- createStyle(fontSize = 9, fontColour = "black")
#' 
#' c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#' c2
#' 
#' writeComment(wb, 1, col = 6 , row = 3, comment = c2)
#' 
#' saveWorkbook(wb, file = "createCommentExample.xlsx", overwrite = TRUE)
createComment <- function(comment,
                          author = Sys.getenv("USERNAME"),
                          style = NULL,
                          width = 2,
                          height = 4){
  
  
  
  if(!"character" %in% class(author))
    stop("author argument must be a character vector")
  
  if(!"character" %in% class(comment))
    stop("comment argument must be a character vector")
  
  if(!"numeric" %in% class(width))
    stop("width argument must be a numeric vector")
  
  if(!"numeric" %in% class(height))
    stop("height argument must be a numeric vector")
  
  
  width <- round(width)
  height <- round(height)
  
  n <- length(comment)
  author <- author[[1]]
  
  if(is.null(style))
    style <- createStyle(fontName = "Tahoma", fontSize = 9, fontColour = "black")
  
  author <- replaceIllegalCharacters(author)
  comment <- replaceIllegalCharacters(comment)
  
  
  invisible(Comment$new(text = comment, author = author, style = style, width = width[1], height = height[1]))
  
}





#' @name writeComment
#' @title write a cell comment
#' @param wb A workbook object
#' @param sheet A vector of names or indices of worksheets
#' @param col Column a column number of letter 
#' @param row A row number.
#' @param comment A Comment object. See \code{\link{createComment}}.
#' @param xy An alternative to specifying \code{col} and
#' \code{row} individually.  A vector of the form
#' \code{c(col, row)}.
#' @export
#' @seealso \code{\link{createComment}}
#' @examples
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet 1")
#' 
#' c1 <- createComment(comment = "this is comment")
#' writeComment(wb, 1, col = "B", row = 10, comment = c1)
#' 
#' s1 <- createStyle(fontSize = 12, fontColour = "red", textDecoration = c("BOLD"))
#' s2 <- createStyle(fontSize = 9, fontColour = "black")
#' 
#' c2 <- createComment(comment = c("This Part Bold red\n\n", "This part black"), style = c(s1, s2))
#' c2
#' 
#' writeComment(wb, 1, col = 6 , row = 3, comment = c2)
#' 
#' saveWorkbook(wb, file = "writeCommentExample.xlsx", overwrite = TRUE)
writeComment <- function(wb, sheet, col, row, comment, xy = NULL){
  
  if(!"Workbook" %in% class(wb))
    stop("First argument must be a Workbook.")
  
  if(!"Comment" %in% class(comment))
    stop("comment argument must be a Comment object")
  
  
  if(length(comment$style) == 1){
    rPr <- wb$createFontNode(comment$style)
  }else{
    rPr <- sapply(comment$style, function(x) wb$createFontNode(x))
  }
  
  rPr <- gsub("font>", "rPr>", rPr)
  sheet <- wb$validateSheet(sheet)
  
  ## All input conversions/validations
  if(!is.null(xy)){
    if(length(xy) != 2)
      stop("xy parameter must have length 2")
    col <- xy[[1]]
    row <- xy[[2]]
  }
  
  if(!is.numeric(col))
    col <- convertFromExcelRef(col)
  
  ref <- paste0(.Call("openxlsx_convert2ExcelRef", col, LETTERS), row)
  
  comment_list <- list("ref" = ref,
                       "author" = comment$author,
                       "comment" = comment$text,
                       "style" = rPr,
                       "clientData" = genClientData(col, row, height = comment$height, width = comment$width))
  
  wb$comments[[sheet]] <- append(wb$comments[[sheet]], list(comment_list))
  
  invisible(wb)
  
}


