#' @title make_oracle_cxn
#' @description This function contains the logic for creating a connection to
#' Oracle.
#' @param fn.oracle.username default is \code{'_none_'} This is your username for
#' accessing oracle objects. If you have a value for this stored in your
#' environment (e.g. from an rprofile file), this can be left and that value will
#' be used.  If a value for this is provided, it will take priority over your
#' existing value.
#' @param fn.oracle.password default is \code{'_none_'} This is your password for
#' accessing oracle objects. If you have a value for this stored in your
#' environment (e.g. from an rprofile file), this can be left and that value will
#' be used.  If a value for this is provided, it will take priority over your
#' existing value.
#' @param fn.oracle.dsn default is \code{'_none_'} This is your dsn/ODBC
#' identifier for accessing oracle objects. If you have a value for this stored in your
#' environment (e.g. from an rprofile file), this can be left and that value will
#' be used.  If a value for this is provided, it will take priority over your
#' existing value.
#' @family data
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @importFrom RODBC odbcConnect
#' @export
#' @note This is stolen directly from the get_data() function of bio.datawrangling
#' (also created by me), but references to roracle were removed
make_oracle_cxn <- function(fn.oracle.username ="_none_",
                            fn.oracle.password="_none_",
                            fn.oracle.dsn="_none_") {
  # Prompt for choice if no db selected ------------------------------------
  if (fn.oracle.username != "_none_"){
    oracle.username= fn.oracle.username
  }
  if (fn.oracle.password != "_none_"){
    oracle.password= fn.oracle.password
  }
  if (fn.oracle.dsn != "_none_"){
    oracle.dsn= fn.oracle.dsn
  }

  use.rodbc <-function(oracle.dsn, oracle.username, oracle.password){
    #warning supressed to hide password should the connection fail
    suppressWarnings(
      assign('oracle_cxn', odbcConnect(oracle.dsn, uid = oracle.username, pwd = oracle.password, believeNRows = F))
    )
    if (class(oracle_cxn) == "RODBC") {
      #cat("\nSuccessfully connected to Oracle via RODBC")
      results = list('rodbc', oracle_cxn)
      return(results)
    } else {
      cat("\nRODBC connection attempt failed\n")
      return(-1)
    }
  }
    #get connection info - only prompt for values not in rprofile
    if (exists('oracle.username')){
      oracle.username <- oracle.username
      #cat("\nUsing stored 'oracle.username'")
    }else{
      oracle.username <-
        readline(prompt = "Oracle Username: ")
      print(oracle.username)
    }

    if (exists('oracle.password')){
      oracle.password <- oracle.password
      #cat("\nUsing stored 'oracle.password'")
    } else {
      oracle.password <-
        readline(prompt = "Oracle Password: ")
      print(oracle.password)
    }

    if (exists('oracle.dsn')){
      oracle.dsn <- oracle.dsn
      #cat("\nUsing stored 'oracle.dsn'")
    }else{
      oracle.dsn <-
        readline(prompt = "Oracle DSN (e.g. PTRAN): ")
      print(oracle.dsn)
    }
      use.rodbc(oracle.dsn, oracle.username, oracle.password)
  }
