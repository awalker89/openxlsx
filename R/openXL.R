#' @name openXL
#' @title Open an xlsx file
#' @author Luca Braglia
#' @description This function tries to open an xlsx file with the
#' proper application, in a portable manner.
#' @param file path to the Excel (xls/xlsx) file
#' @usage openXL(file=NULL)
#' @export openXL
#' @examples
#' example(writeData)
#' openXL("writeDataExample.xlsx")
#' 
openXL <- function(file = NULL){

    if (is.null(file)) stop("a file have to be specified")

    ## execution should be in background in order to not block R
    ## interpreter
    this.system <- Sys.info()["sysname"]
    if ("Linux" == this.system ) {
        if (is.null(app <- unlist(options('openxlsx.excel.app')))) {
            app <- chooseExcelApp()
        }
        my.command <- paste(app, file, "&", sep = " ")
        system(command = my.command)
    } else if ("Windows" == this.system ){
        shell(shQuote(string = file)) 
    } else if ("Darwin" == this.system){
        my.command <- paste0("open ", file)
        system(command = my.command)
    } else {
        warning("Operative system not handled.")
    }
    
}



#' @name chooseExcelApp
#' @title Search for available xls/xlsx reader application in a
#' UNIX like (mainly Linux) environment. 
#' @author Luca Braglia
#' @description This function search for available xls/xlsx
#' reader application. If it founds anything, sets
#' \code{options('openxlsx.excel.app')}, used by \code{openXL}.
#' @usage chooseExcelApp()
#' @export chooseExcelApp
chooseExcelApp <- function() {

    if ("Linux" != (this.system <- Sys.info()["sysname"] ))
        stop("Currently you should't need to run chooseExcelApp on ",
             this.system,". Try using openXL directly.")
    
    
    find_xl_bin <- function(program) {
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
           `Gnumeric` = "gnumeric"
            )

    prog <- sapply(m, find_xl_bin)
    n.apps <- length(avail.prog <- prog[ "Not Available" != prog])

    if (0 == n.apps) {
        stop("No application (detected) availables.\n",
             "Set options('openxlsx.excel.app'), instead." )
    } else if (1 == n.apps) {
        cat("Only ", names(avail.prog), "found; I'll use it.\n")
        unnprog <- unname(avail.prog)
        options(openxlsx.excel.app = unnprog)
        invisible(unnprog)
    } else if (1 < n.apps) {
        if (!interactive())
            stop("Cannot choose an Excel file opener non-interactively.\n",
                 "Set options('openxlsx.excel.app'), instead.")
        res <- menu(names(avail.prog), title = "Excel Apps availables")
        unnprog <- unname(avail.prog[res])
        if (res > 0L) options(openxlsx.excel.app = unnprog)
        invisible(unname(unnprog))
    } else {
        stop("Unexpected error")
    }
}
