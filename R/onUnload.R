.onUnload <- function (libpath) {
  library.dynam.unload("openxlsx", libpath)
}