#' @name writeData
#' @title Write an object to a worksheet
#' @author Alexander Walker
#' @description Write an object to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Object to be written.
#' @param startCol A vector specifiying the starting columns(s) to write df
#' @param startRow A vector specifiying the starting row(s) to write df
#' @param xy An alternative to specifying startCol and startRow individually.
#'  A vector of the form c(startCol, startRow)
#' @param colNames If TRUE, column names of x are written.
#' @param rowNames If TRUE, data.frame row names of x are written.
#' @param headerStyle Custom style to apply to column names.
#' @param borders Either "none" (default), "surrounding",
#' "columns", "rows" or respective abbreviations.  If
#' "surrounding", a border is drawn around the data.  If "rows",
#' a surrounding border is drawn a border around each row. If
#' "columns", a surrounding border is drawn with a border between
#' each column.
#' @param borderColour Colour of cell border.  A valid colour (belonging to colours() or a hex colour code.)
#' @param ...  Further arguments (for future use)
#' @seealso \code{\link{writeData}}
#' @export writeData
#' @rdname writeData
#' @examples
#' 
#' \dontrun{
#' ## inspired by xtable gallery
#' ##' http://cran.r-project.org/web/packages/xtable/vignettes/xtableGallery.pdf
#' 
#' ## Create a new workbook and delete old file, if existing
#' wb <- createWorkbook()
#' my.file <- "test.xlsx"
#' unlink(my.file)
#' data(tli, package = "xtable")
#' 
#' ## TEST 1 - data.frame
#' test.n <- "data.frame"
#' my.df <- tli[1:10, ]
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.df, borders = "n")
#' 
#' ## TEST 2 - matrix
#' test.n <- "matrix"
#' design.matrix <- model.matrix(~ sex * grade, data = my.df)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = design.matrix)
#' 
#' ## TEST 3 - aov
#' test.n <- "aov"
#' fm1 <- aov(tlimth ~ sex + ethnicty + grade + disadvg, data = tli)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm1)
#' 
#' ## TEST 4 - lm
#' test.n <- "lm"
#' fm2 <- lm(tlimth ~ sex*ethnicty, data = tli)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm2)
#' 
#' ## TEST 5 - anova 1 
#' test.n <- "anova"
#' my.anova <- anova(fm2)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.anova)
#' 
#' ## TEST 6 - anova 2
#' test.n <- "anova2"
#' fm2b <- lm(tlimth ~ ethnicty, data = tli)
#' my.anova2 <- anova(fm2b, fm2)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.anova2)
#' 
#' ## TEST 7 - GLM
#' test.n <- "glm"
#' fm3 <- glm(disadvg ~ ethnicty*grade, data = tli, family =
#'            binomial())
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = fm3)
#' 
#' ## TEST 8 - prcomp
#' test.n <- "prcomp"
#' pr1 <- prcomp(USArrests)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = pr1)
#'
#' ## TEST 9 - summary.prcomp
#' test.n <- "summary.prcomp"
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = summary(pr1))
#'
#' ## TEST 10 - univariate ts
#' test.n <- "u.ts"
#' set.seed(1)
#' uts <- ts(cumsum(1+round(rnorm(100), 0)),
#'           start = c(1954,7),
#'           frequency = 12)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = uts)
#'
#' ## TEST 11 - multivariate ts (from help(ts))
#' test.n <- "m.ts"
#' mts <- ts(matrix(rnorm(300), 100, 3), start = c(1961, 1), frequency = 12)
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = mts)
#'
#' ## TEST 12 - univariate ts (not calendar like)
#' test.n <- "u.ts.2" 
#' (uts2 <- ts(cumsum(1+round(rnorm(100), 0)),
#'           start = c(1954,7),
#'           frequency = 6))
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = uts2)
#' 
#' ## TEST 13 - univariate ts (calendar like, obs <= freq)
#' test.n <- "u.ts.3" 
#' (uts3 <- ts(cumsum(1+round(rnorm(2), 0)),
#'           start = c(1954,1),
#'           frequency = 4))
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = uts3)
#'
#' ## TEST 14 - univariate ts (yearly series, obs <= freq)
#' test.n <- "u.ts.4" 
#' (uts4 <- ts(cumsum(1+round(rnorm(2), 0)),
#'           start = c(1954,1),
#'           frequency = 1))
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = uts4)
#' 
#' ## TEST XX - simple table
#' test.n <- "table"
#' data(airquality)
#' airquality$OzoneG80 <- factor(airquality$Ozone > 80,
#'                               levels = c(FALSE, TRUE),
#'                               labels = c("Oz <= 80", "Oz > 80"))
#' airquality$Month <- factor(airquality$Month,
#'                            levels = 5:9,
#'                            labels = month.abb[5:9])
#' my.table <- with(airquality, table(OzoneG80,Month) )
#' addWorksheet(wb = wb, sheetName = test.n)
#' writeData(wb = wb, sheet = test.n, x = my.table)
#' 
#' ## Save workbook
#' saveWorkbook(wb, my.file,  overwrite = TRUE)
#' }
writeData <- function(wb, 
                      sheet,
                      x,
                      startCol = 1,
                      startRow = 1, 
                      xy = NULL,
                      colNames = TRUE,
                      rowNames = FALSE,
                      headerStyle = NULL,
                      borders = c("none","surrounding","rows","columns"),
                      borderColour = getOption("openxlsx.borderColour", "black"),
                      ...){
  
  ## All input conversions/validations
  if(!is.null(xy)){
    if(length(xy) != 2)
      stop("xy parameter must have length 2")
    startCol <-  xy[[1]]
    startRow <- xy[[2]]
  }
  
  ## convert startRow and startCol
  startCol <- convertFromExcelRef(startCol)
  startRow <- as.integer(startRow)
  
  if(!"Workbook" %in% class(wb)) stop("First argument must be a Workbook.")
  if(!is.logical(colNames)) stop("colNames must be a logical.")
  if(!is.logical(rowNames)) stop("rowNames must be a logical.")
  if(!is.null(headerStyle) & !"Style" %in% class(headerStyle)) stop("headerStyle must be a style object or NULL.")
  
  borders <- match.arg(borders)
  if(length(borders) != 1) stop("borders argument must be length 1.")    
  
  ## borderColours validation
  borderColour <- validateBorderColour(borderColour)
  
  clx <- class(x)
  Call <- match.call()
  xname <- paste(Call$x)
  
  if(any(c("data.frame", "data.table") %in% clx)){
    ## Do nothing
  }else if("matrix" %in% clx){
    x <- as.data.frame(x, stringsAsFactors = FALSE)
    
  }else if("array" %in% clx){
    stop("array in writeData : currently not supported")
    
  }else if("aov" %in% clx){
    
    x <- summary(x)
    x <- cbind(x[[1]])
    x <- cbind(data.frame("row name" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE 
    
  }else if("lm" %in% clx){
    
    x <- as.data.frame(summary(x)[["coefficients"]])
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
        
  }else if("anova" %in% clx){
    
    x <- cbind(x)
    x <- cbind(data.frame("row name" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("glm" %in% clx){
    
    x <- as.data.frame(summary(x)[["coefficients"]])
    x <- cbind(data.frame("row name" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("table" %in% clx){
      
    x <- as.data.frame(unclass(x))
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("prcomp" %in% clx){
    
    x <- as.data.frame(x$rotation)
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE
    
  }else if("summary.prcomp" %in% clx){
    
    x <- as.data.frame(x$importance)
    x <- cbind(data.frame("Variable" = rownames(x)), x)
    names(x)[1] <- ""
    rowNames <- FALSE

  }else if("mts" %in% clx[1]){ ## Multivariate ts
      ## HELP TODO BUG HERE: why doesn't this match (and it
      ## goes to else?)... I can't figure it out, for the moment :(
   
      x <- as.data.frame(as.matrix(x))
      rowNames <- TRUE
      colNames <- TRUE

  }else if("ts" %in% clx[1]){           #univariate ts
      ## simplified/little changed version of print.ts with .preformat.ts
      Tsp <- tsp(x)
      if (is.null(Tsp))  stop("series is corrupt, with no 'tsp' attribute")
      fr.x <- frequency(x)
      ## do calendar iff a monthly/quaterly series
      calendar <- any(fr.x == c(4, 12)) && length(start(x)) == 2L
      if (calendar) { ##do calendar-like
          rowNames <- TRUE
          colNames <- TRUE
          if (fr.x > 1) {                 
              dn2 <- if (fr.x == 12)      #monthly series
                  month.abb
              else if (fr.x == 4) {       #quarterly series
                  c("Qtr1", "Qtr2", "Qtr3", "Qtr4")
              } #else paste0("p", 1L:fr.x)
                #TODO: other"ly" series (not calendar now)

              ## the series starts and ends in the same year with a
              ## number of obs <= its frequency
              if (NROW(x) <= fr.x && start(x)[1L] == end(x)[1L]) {
                  ## rows
                  dn1 <- start(x)[1L]
                  ## cols
                  dn2 <- dn2[1 + (start(x)[2L] - 2 + seq_along(x))%%fr.x]
                  attributes(x) <- NULL
                  x <- as.data.frame(matrix(x, nrow = 1L, byrow = TRUE,
                                            dimnames = list(dn1, dn2)))
              } else { ##otherwise (eg start and end in different years)
                  start.pad <- start(x)[2L] - 1 # n of initial blank
                  end.pad <- fr.x - end(x)[2L] # last blank
                  dn1 <- start(x)[1L]:end(x)[1L] # rowNames
                  attributes(x) <- NULL # makes it a vector
                  vec <- c(rep.int(NA, start.pad), x,
                           rep.int(NA, end.pad))
                  x <- as.data.frame(matrix(vec, ncol = fr.x,  
                                            byrow = TRUE,
                                            dimnames = list(dn1, dn2)))
              }

          } else { ##(n-)yearly series
              times <- unlist(lapply(time(x), paste)) # placeholder?
              attributes(x) <- NULL
              x <- data.frame(x)
              row.names(x) <- times
              names(x) <- xname
              rowNames <- TRUE
              colNames <- TRUE
          }
      } else { ## not calendar-like
          times <- unlist(lapply(time(x), paste)) # placeholder?
          attr(x, "class") <- attr(x, "tsp") <- attr(x, "na.action") <- NULL
          x <- as.data.frame(x)
          row.names(x) <- times
          names(x) <- xname
          rowNames <- TRUE
          colNames <- TRUE
      }
            
  }else{
      ## TODO check if as.data.frame method exists, otherwise stop()
      x <- as.data.frame(x, stringsAsFactors = FALSE)
      colNames <- FALSE
      rowNames <- FALSE
  }
  
  
  ## cbind rownames to x
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x)), x)
    names(x)[[1]] <- ""
  }
  
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style  
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             rows = startRow,
             cols = 0:(nCol-1) + startCol,
             gridExpand = TRUE)
  
  ## draw borders
  if( "none" != borders )
    doBorders(borders = borders, wb = wb,
              sheet = sheet, startRow = startRow + colNames,
              startCol = startCol, nrow = nrow(x), ncol = ncol(x),
              borderColour = borderColour)
  
}
