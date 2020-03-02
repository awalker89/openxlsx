

##

openxlsxCoerce <- function(x, rowNames){
  
  UseMethod("openxlsxCoerce") 
  
}

openxlsxCoerce.default <- function(x, rowNames){

  x <- as.data.frame(x, stringsAsFactors = FALSE)
  return(x)

}


openxlsxCoerce.data.frame <- function(x, rowNames){

  ## cbind rownames to x
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x), stringsAsFactors = FALSE), as.data.frame(x, stringsAsFactors = FALSE))
    names(x)[[1]] <- ""
  }
  
  return(x)
  
}


openxlsxCoerce.data.table <- function(x, rowNames){

  x <- as.data.frame(x, stringsAsFactors = FALSE)
  
  ## cbind rownames to x
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x), stringsAsFactors = FALSE), x)
    names(x)[[1]] <- ""
  }
  
  return(x)
  
}


openxlsxCoerce.matrix <- function(x, rowNames){
  
  x <- as.data.frame(x, stringsAsFactors = FALSE)
  
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x), stringsAsFactors = FALSE), x)
    names(x)[[1]] <- ""
  }
  
  return(x)
  
}


openxlsxCoerce.array <- function(x, rowNames){
  stop("array in writeData : currently not supported")
}

openxlsxCoerce.aov <- function(x, rowNames){
  
  x <- summary(x)
  x <- cbind(x[[1]])
  x <- cbind(data.frame("row name" = rownames(x), stringsAsFactors = FALSE), x)
  names(x)[1] <- ""
  
  return(x)
}


openxlsxCoerce.lm <- function(x, rowNames){
  
  x <- as.data.frame(summary(x)[["coefficients"]])
  x <- cbind(data.frame("Variable" = rownames(x), stringsAsFactors = FALSE), x)
  names(x)[1] <- ""
  
  return(x)
}


openxlsxCoerce.anova <- function(x, rowNames){
  
  x <- as.data.frame(x)
  
  if(rowNames){
    x <- cbind(data.frame("row name" = rownames(x), stringsAsFactors = FALSE), x)
    names(x)[1] <- ""
  }
  
  return(x)
}


openxlsxCoerce.glm <- function(x, rowNames){
  
  x <- as.data.frame(summary(x)[["coefficients"]])
  x <- cbind(data.frame("row name" = rownames(x), stringsAsFactors = FALSE), x)
  names(x)[1] <- ""
  
  return(x)
}


openxlsxCoerce.table <- function(x, rowNames){
  
  x <- as.data.frame(unclass(x))
  x <- cbind(data.frame("Variable" = rownames(x), stringsAsFactors = FALSE), x)
  names(x)[1] <- ""
  
  return(x)
}


openxlsxCoerce.prcomp <- function(x, rowNames){
  
  x <- as.data.frame(x$rotation)
  x <- cbind(data.frame("Variable" = rownames(x), stringsAsFactors = FALSE), x)
  names(x)[1] <- ""
  
  return(x)
}


openxlsxCoerce.summary.prcomp <- function(x, rowNames){
  
  x <- as.data.frame(x$importance)
  x <- cbind(data.frame("Variable" = rownames(x), stringsAsFactors = FALSE), x)
  names(x)[1] <- ""
  
  return(x)
}



openxlsxCoerce.survdiff <- function(x, rowNames){
  
  
  ## like print.survdiff with some ideas from the ascii package
  if(length(x$n) == 1){
    
    z <- sign(x$exp - x$obs) * sqrt(x$chisq)
    temp <- c(x$obs, x$exp, z, 1 - pchisq(x$chisq,  1))
    names(temp) <- c("Observed", "Expected", "Z", "p")
    x <- as.data.frame(t(temp))
  
  }else{  
    
    if(is.matrix(x$obs)) {
      otmp <- apply(x$obs, 1, sum)
      etmp <- apply(x$exp, 1, sum)
    }
    else {
      otmp <- x$obs
      etmp <- x$exp
    }
    chisq <- c(x$chisq, rep(NA, length(x$n) - 1))
    df <- c((sum(1 * (etmp > 0))) - 1, rep(NA, length(x$n) - 1))
    p <- c(1 - pchisq(x$chisq, df[!is.na(df)]), rep(NA, length(x$n) - 1))
    
    temp <- cbind(x$n, otmp, etmp, 
                  ((otmp - etmp)^2)/etmp, ((otmp - etmp)^2)/diag(x$var),
                  chisq, df, p)
    
    
    colnames(temp) <- c("N", "Observed", "Expected", "(O-E)^2/E", "(O-E)^2/V",
                        "Chisq", "df","p")
    
    temp <- as.data.frame(temp, checknames = FALSE)
    x <- cbind("Group" = names(x$n), temp)
    names(x)[1] <- ""
    
  }

  return(x)
}



openxlsxCoerce.coxph <- function(x, rowNames){
  
  ## sligthly modified print.coxph
  coef <- x$coefficients
  se <- sqrt(diag(x$var))
  
  if(is.null(coef) | is.null(se)) 
    stop("Input is not valid")
  
  if(is.null(x$naive.var)){
    tmp <- cbind(coef, exp(coef), se, coef/se, pchisq((coef/se)^2, 1))
    colnames(tmp) <- c("coef", "exp(coef)", "se(coef)", "z", "p")
  
  }else{
    nse <- sqrt(diag(x$naive.var))
    tmp <- cbind(coef, exp(coef), nse, se, coef/se, pchisq((coef/se)^2, 1))
    colnames(tmp) <- c("coef", "exp(coef)", "se(coef)", "robust se", "z", "p")
  }
  
  x <- cbind("Variable" = names(coef), as.data.frame(tmp, checknames = FALSE))
  names(x)[1] <- ""

  return(x)
}




openxlsxCoerce.summary.coxph <- function(x, rowNames){
  
  coef <- x$coefficients
  ci <- x$conf.int
  nvars <- nrow(coef)
  
  tmp <- cbind(coef[, - ncol(coef), drop=FALSE],          #p later
               ci[, (ncol(ci) - 1):ncol(ci), drop=FALSE], #confint
               coef[, ncol(coef), drop=FALSE])            #p.value
  
  x <- as.data.frame(tmp, checknames = FALSE)
  
  x <- cbind(data.frame("row names" = rownames(x)), x)
  names(x)[[1]] <- ""
  
  return(x)
  
}

openxlsxCoerce.cox.zph <- function(x, rowNames){
  
  tmp <- as.data.frame(x$table)
  x <- cbind(data.frame("row names" = rownames(tmp)), tmp)
  names(x)[[1]] <- ""
  
  return(x)
  
}


openxlsxCoerce.hyperlink <- function(x, rowNames){
  
  ## vector of hyperlinks
  class(x) <- c("character", "hyperlink") 
  x <- as.data.frame(x, stringsAsFactors = FALSE)
  
}



