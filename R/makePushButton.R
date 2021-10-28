#' @title makePushButton
#' @description This function facilitates the generation of a batch file to your
#' desktop that can be double-clicked to run reports easily
#' @param arch default is \code{NULL}.  This is the r architecture you want to use - either 32 or 64 are 
#' valid entries.  If left as NULL, it will use your default architecture
#' @family Mar.portsampling
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @importFrom RODBC odbcConnect
#' @export
makePushButton<-function(arch = NULL){
  if (Sys.info()["sysname"]=="Linux"){
    stop("No batch files for Linux - Sorry!")
  }else{
    thePath  = file.path("C:","Users",Sys.info()["user"],"Desktop")
  }

  thePath =  path.expand(thePath)
  dir.create(thePath, showWarnings = FALSE)
  filename = "runPortSamplingRpt.bat"
  batFile = file(file.path(thePath,filename))
  RSLoc = file.path(R.home("bin"),"Rscript.exe")
  if (is.null(arch)){
  }else if (arch ==32){
    RSLoc <- gsub(pattern = "x64", replacement = "i386", RSLoc)
  }else{
    RSLoc <- gsub(pattern = "i386", replacement = "x64", RSLoc)
  }
  
  dir.create(file.path("C:","DFO-MPO","PORTSAMPLING"), showWarnings = FALSE)
  runScript = file.path("C:","DFO-MPO","PORTSAMPLING","batchRun.R")


  
  oracle.username <- readline(prompt = "Oracle Username: ")
  print(oracle.username)

  oracle.password <- readline(prompt = "Oracle Password: ")
  print(oracle.password)

  oracle.dsn <- readline(prompt = "Oracle DSN (e.g. PTRAN): ")
  print(oracle.dsn)

  usepkg <- readline(prompt = "rodbc or roracle?  (proably rodbc): ")
  print(usepkg)

  head="REM PortSampling Report Runner
REM Double-clicking this file will run your reports automatically
REM
REM You can generate a file like this automatically via
REM Mar.portsampling::makePortSamplingBat()
REM"
  scriptPath = paste0('"',RSLoc,'" "',runScript,'" ', oracle.username, ' ', oracle.password, ' ', oracle.dsn, ' ', usepkg)
  writeLines(c(head, scriptPath,'PAUSE'), batFile)
  close(batFile)

  filename2 = "batchRun.R"
  runFile = file(file.path("C:","DFO-MPO","PORTSAMPLING",filename2))

runner = "batchRun<-function(args){
    Mar.portsampling::makeHailInRpt(thePath = 'C:/DFO-MPO/PORTSAMPLING',
                  fn.oracle.username = args[1],
                  fn.oracle.password = args[2],
                  fn.oracle.dsn = args[3],
                  usepkg = args[4])
}
batchRun(commandArgs(TRUE))
"
writeLines(runner, runFile)


  cat(paste0("Batch file written to ",file.path(thePath,filename)))
}
