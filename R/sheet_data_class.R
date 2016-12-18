

#' @include class_definitions.R

Sheet_Data$methods(initialize = function(){
  
  rows <<- integer(0)
  cols <<- integer(0)
  
  t <<- integer(0)
  v <<- character(0)
  f <<- character(0)
  
  style_id <<- character(0)
  
  data_count <<- 0L
  n_elements <<- 0L
  
})



Sheet_Data$methods(delete = function(row_in, col_in, grid_expand){
  
  col_in <- convertFromExcelRef(col_in)
  row_in <- as.integer(row_in)
  
  ## rows and cols need to be the same length
  if(grid_expand){
    combs <- expand.grid(row_in, col_in) 
    row_in <- combs[,1]
    col_in <- combs[,2]
  }
  
  if(length(row_in) != length(col_in)){
    stop("Length of rows and cols must be equal.")
  }
  
  
  inds <- which(paste(rows, cols, sep = ",") %in% paste(row_in, col_in, sep = ","))
  
  if(length(inds) > 0){ ## writing over existing data 
    
    rows <<- rows[-inds]
    cols <<- cols[-inds]
    t    <<- t[-inds]
    v    <<- v[-inds]
    f    <<- f[-inds]
  
    n_elements <<- as.integer(length(rows))

    if(n_elements == 0)
      data_count <<- 0L
  }
  
})



Sheet_Data$methods(write = function(row_in, col_in, t_in, v_in, f_in){
  
  coords <- expand.grid(col_in, row_in)
  col_in <- coords[[1]]
  row_in <- coords[[2]]
  
  v_in[!is.na(f_in)] <- as.character(NA)
  t_in[!is.na(f_in)] <- 3L ## "str"
  
  if(length(row_in) != 0 & length(col_in) != 0){
    
    if(data_count == 0 | n_elements == 0){
      
      rows <<- row_in
      cols <<- col_in
      t <<- t_in
      v <<- v_in
      f <<- f_in

    }else{
      
      inds <- which(paste(rows, cols, sep = ",") %in% paste(row_in, col_in, sep = ","))
      
      if(length(inds) > 0){ ## writing over existing data 
        
        rows <<- c(rows[-inds], row_in)
        cols <<- c(cols[-inds], col_in)
        t <<- c(t[-inds], t_in)
        v <<- c(v[-inds], v_in)
        f <<- c(f[-inds], f_in)
        
        
      }else{ ## no overlap
        
        rows <<- c(rows, row_in)
        cols <<- c(cols, col_in)
        t <<- c(t, t_in)
        v <<- c(v, v_in)
        f <<- c(f, f_in)
        
      }
      
      
    } ## make sure writing some data
  }
  
  n_elements <<- as.integer(length(rows))
  data_count <<- data_count + 1L
  
  
})
