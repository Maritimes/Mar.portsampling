#' @title update.portsampling
#' @description This is just a quick and dirty script for facilitating updates
#' @family portsampling
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @importFrom devtools install_github
#' @importFrom utils remove.packages
#' @export
update.portsampling <- function(){
  remove.packages(pkgs = "bio.portsampling")
  install_github('Maritimes/bio.portsampling')
}
