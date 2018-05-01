.onAttach <- function(libname, pkgname) {
  packageStartupMessage(paste0("###\nVersion: ", asNamespace("Mar.portsampling")$'.__NAMESPACE__.'$spec[['version']]))
}
.onLoad <- function(libname, pkgname){
  options(stringsAsFactors = FALSE)
 # assign("ds_all", load_datasources(), envir = .GlobalEnv)
}
