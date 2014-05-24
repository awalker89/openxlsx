#' @name openXL
#' @title Open a Microsoft Excel file (xls/xlsx) or an openxlsx Workbook
#' @author Luca Braglia
#' @description This function tries to open a Microsoft Excel
#' (xls/xlsx) file or an openxlsx Workbook with the proper
#' application, in a portable manner.
#'
#' In Linux it searches for available xls/xlsx reader application
#' (unless \code{options('openxlsx.excelApp')} is set to the bin
#' app path), and if it founds anything, sets
#' \code{options('openxlsx.excelApp')}.
#' 
#' @param file path to the Excel (xls/xlsx) file or Workbook object.
#' @usage openXL(file=NULL)
#' @export openXL
#' @examples
#' # file example
#' example(writeData)
#' openXL("writeDataExample.xlsx")
#'
#' # (not yet saved) Workbook example
#' wb <- createWorkbook()
#' x <- mtcars[1:6,]
#' addWorksheet(wb, "Cars")
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' openXL(wb)
#' 


openXL <- function(file = NULL){

    if (is.null(file)) stop("a file have to be specified")

    ## workbook handling
    if ("Workbook" == class(file)) {
        oldWD <- getwd()
        file <- file.path(file$saveWorkbook(quiet = TRUE), "temp.xlsx")
        setwd(oldWD)
    }
    
    if (!file.exists(file)) stop("Non existent file or wrong path.")
    
    ## execution should be in background in order to not block R
    ## interpreter
    userSystem <- Sys.info()["sysname"]

    
    if ("Linux" == userSystem ) {
        if (is.null(app <- unlist(options('openxlsx.excelApp')))) {
            app <- chooseExcelApp()
        }
        myCommand <- paste(app, file, "&", sep = " ")
        system(command = myCommand)
        
    } else if ("Windows" == userSystem ){
        shell(shQuote(string = file))
        
    } else if ("Darwin" == userSystem){
        myCommand <- paste0("open ", file)
        system(command = myCommand)
        
    } else {
        warning("Operative system not handled.")
    }
    
}


## chooseExcelApp -- start
chooseExcelApp <- function() {
    
    findXLbin <- function(program) {
        con <- pipe(paste0("which ",program))
        res <- readLines(con, n=1)
        close(con)

        if (0 != length(res)){
            return(res)
        } else {
            return("Not Available")
        }
    }

    ## TODO: other useful executables to search for?
    m <- c(`Libreoffice/OpenOffice` = "soffice",
           `Gnumeric` = "gnumeric")

    prog <- sapply(m, findXLbin)
    nApps <- length(availProg <- prog[ "Not Available" != prog])

    if (0 == nApps) {
        stop("No application (detected) availables.\n",
             "Set options('openxlsx.excelApp'), instead." )
    } else if (1 == nApps) {
        cat("Only ", names(availProg), "found; I'll use it.\n")
        unnprog <- unname(availProg)
        options(openxlsx.excelApp = unnprog)
        invisible(unnprog)
    } else if (1 < nApps) {
        if (!interactive())
            stop("Cannot choose an Excel file opener non-interactively.\n",
                 "Set options('openxlsx.excelApp'), instead.")
        res <- menu(names(availProg), title = "Excel Apps availables")
        unnprog <- unname(availProg[res])
        if (res > 0L) options(openxlsx.excelApp = unnprog)
        invisible(unname(unnprog))
    } else {
        stop("Unexpected error")
    }
}
