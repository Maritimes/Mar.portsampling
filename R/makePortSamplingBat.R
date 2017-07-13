#' @title makePortSamplingBat
#' @description This function facilitates the generation of a batch file to your
#' desktop that can be double-clicked to run reports easily
#' @family portsampling
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @importFrom RODBC odbcConnect
#' @export
#' @note This is stolen directly from the get_data() function of bio.datawrangling
#' (also created by me), but references to roracle were removed
makePortSamplingBat<-function(){

  if (Sys.info()["sysname"]=="Linux"){
    stop("No batch files for Linux - Sorry!")
  }else{
    thePath  = file.path("C:","Users",Sys.info()["user"],"Desktop")
  }

  thePath =  path.expand(thePath)
  dir.create(thePath, showWarnings = FALSE)
  batFile = file(file.path(thePath,"runPortSamplingRpt.bat"))
  RSLoc = file.path(R.home("bin"),"Rscript.exe")
  runScript = file.path(getwd(),"inst","batchRun.R")
  RSLoc = file.path(R.home("bin"),"Rscript.exe")

  oracle.username <- readline(prompt = "Oracle Username: ")
  print(oracle.username)

  oracle.password <- readline(prompt = "Oracle Password: ")
  print(oracle.password)

  oracle.dsn <- readline(prompt = "Oracle DSN (e.g. PTRAN): ")
  print(oracle.dsn)


  head="REM PortSampling Report Runner
REM Double-clicking this file will run your reports automatically
REM
REM You can generate a file like this automatically via
REM bio.portsampling::makePortSamplingBat()
REM"

  scriptPath = paste0('"',RSLoc,'" "',runScript,'" ', oracle.username, ' ', oracle.password, ' ', oracle.dsn)
  writeLines(c(head, scriptPath,'PAUSE'), batFile)
  close(batFile)
  cat(paste0("Batch file written to ",file.path(thePath,"test.bat")))
}
