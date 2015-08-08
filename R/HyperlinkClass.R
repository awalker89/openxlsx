


Hyperlink <- setRefClass("Hyperlink", 
                         
                         fields = c("ref",
                                    "target",
                                    "location",
                                    "display",
                                    "is_external"),
                         
                         methods = list()
)


Hyperlink$methods(initialize = function(ref, target, location, display = NULL, is_external = TRUE){
  
  ref <<- ref
  target <<- target
  location <<- location
  display <<- display
  is_external <<- is_external
  
})

Hyperlink$methods(to_xml = function(id){
  

  loc <- sprintf('location="%s"', location)
  disp <- sprintf('display="%s"', display)
  rf <- sprintf('ref="%s"', ref)
  
  if(is_external){
    rid <- sprintf('r:id="rId%s"', id)
  }else{
    rid <- NULL
  }
  
  paste('<hyperlink', rf, rid, disp, loc, '/>')
  
})

Hyperlink$methods(to_target_xml = function(id){
  
  
  if(is_external){
    return(sprintf('<Relationship Id="rId%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>', id, target))
  }else{
    return(NULL)
  }

})



xml_to_hyperlink <- function(xml){
 
  # xml <- c('<hyperlink ref="A1" r:id="rId1" location="Authority"/>',
  # '<hyperlink ref="B1" r:id="rId2"/>',
  # '<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/>')
  
  if(length(xml) == 0)
    return(xml)
  
  targets <- names(xml)
  if(is.null(targets))
    targets <- rep(NA, length(xml))
  
  xml <- unname(xml)
  
  a <- unlist(lapply(xml, function(x) regmatches(x, gregexpr('[a-zA-Z]+=".*?"', x))), recursive = FALSE)
  names = lapply(a, function(xml) regmatches(xml, regexpr('[a-zA-Z]+(?=\\=".*?")', xml, perl = TRUE)))
  vals =  lapply(a, function(xml) regmatches(xml, regexpr('(?<=").*?(?=")', xml, perl = TRUE)))
  vals <- lapply(vals, function(x) {Encoding(x) <- "UTF-8"; x})

  lapply(1:length(xml), function(i){
    
    tmp_vals <- vals[[i]]
    tmp_nms <- names[[i]]
    names(tmp_vals) <- tmp_nms
    
    ## ref
    ref <- tmp_vals[["ref"]]
    
    ## location
    if("location" %in% tmp_nms){
      location <- tmp_vals[["location"]]
    }else{
      location <- NULL
    } 
    
    ## location
    if("display" %in% tmp_nms){
      display <- tmp_vals[["display"]]
    }else{
      display <- NULL
    } 
    
    ## target/external
    if(is.na(targets[i])){
      target <- NULL
      is_external <- FALSE
    }else{
      is_external <- TRUE
      target <- targets[i]
    }

    Hyperlink$new(ref = ref, target = target, location = location, display = display, is_external = is_external)
    
  })
  
}







