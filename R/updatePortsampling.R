#' @title updatePortsampling
#' @description This is just a quick and dirty script for facilitating updates
#' @family portsampling
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @importFrom devtools install_github
#' @importFrom utils remove.packages
#' @export
updatePortsampling <- function(){
  if (require(bio.portsampling)) remove.packages(pkgs = "bio.portsampling")
  if (require(portsampling)) remove.packages(pkgs = "portsampling")
  install_github('Maritimes/portsampling')
}
