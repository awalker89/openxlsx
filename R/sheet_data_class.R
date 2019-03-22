

#' @include class_definitions.R

Sheet_Data$methods(initialize = function() {
  rows <<- integer(0)
  cols <<- integer(0)

  t <<- integer(0)
  v <<- character(0)
  f <<- character(0)

  style_id <<- character(0)

  data_count <<- 0L
  n_elements <<- 0L
})



Sheet_Data$methods(delete = function(rows_in, cols_in, grid_expand) {
  cols_in <- convertFromExcelRef(cols_in)
  rows_in <- as.integer(rows_in)

  ## rows and cols need to be the same length
  if (grid_expand) {
    n <- length(rows_in)
    rows_in <- rep.int(rows_in, times = length(cols_in))
    cols_in <- rep(cols_in, each = n)
  }

  if (length(rows_in) != length(cols_in)) {
    stop("Length of rows and cols must be equal.")
  }

  inds <- which(paste(rows, cols, sep = ",") %in% paste(rows_in, cols_in, sep = ","))

  if (length(inds) > 0) { ## writing over existing data

    rows <<- rows[-inds]
    cols <<- cols[-inds]
    t <<- t[-inds]
    v <<- v[-inds]
    f <<- f[-inds]

    n_elements <<- as.integer(length(rows))

    if (n_elements == 0) {
      data_count <<- 0L
    }
  }
})



Sheet_Data$methods(write = function(rows_in, cols_in, t_in, v_in, f_in, any_functions = TRUE) {
  if (length(rows_in) == 0 | length(cols_in) == 0) {
    return(invisible(0))
  }


  possible_overlap <- FALSE
  if (n_elements > 0) {
    possible_overlap <- (min(cols_in, na.rm = TRUE) <= max(cols, na.rm = TRUE)) &
      (max(cols_in, na.rm = TRUE) >= min(cols, na.rm = TRUE)) &
      (min(rows_in, na.rm = TRUE) <= max(rows, na.rm = TRUE)) &
      (max(rows_in, na.rm = TRUE) >= min(rows, na.rm = TRUE))
  }

  n <- length(cols_in)
  cols_in <- rep.int(cols_in, times = length(rows_in))
  rows_in <- rep(rows_in, each = n)

  if (any_functions) {
    if (any(!is.na(f_in))) {
      v_in[!is.na(f_in)] <- as.character(NA)
      t_in[!is.na(f_in)] <- 3L ## "str"
    }
  }

  inds <- integer(0)
  if (possible_overlap) {
    inds <- which(paste(rows, cols, sep = ",") %in% paste(rows_in, cols_in, sep = ","))
  }

  if (length(inds) > 0) {
    rows <<- c(rows[-inds], rows_in)
    cols <<- c(cols[-inds], cols_in)
    t <<- c(t[-inds], t_in)
    v <<- c(v[-inds], v_in)
    f <<- c(f[-inds], f_in)
  } else {
    rows <<- c(rows, rows_in)
    cols <<- c(cols, cols_in)
    t <<- c(t, t_in)
    v <<- c(v, v_in)
    f <<- c(f, f_in)
  }

  n_elements <<- as.integer(length(rows))
  data_count <<- data_count + 1L
})
