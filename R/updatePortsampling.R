#' @title updatePortsampling
#' @description This is just a quick and dirty script for facilitating updates
#' @family Mar.portsampling
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @importFrom devtools install_github
#' @importFrom utils remove.packages
#' @export
updatePortsampling <- function(){
  if (require(Mar.portsampling)) remove.packages(pkgs = "Mar.portsampling")
  install_github('Maritimes/Mar.portsampling')
}
