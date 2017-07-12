batchRun <-function(thePath = file.path("C:","DFO-MPO","PORTSAMPLING"),
                    fn.oracle.username = "<your_oracle_username>",
                    fn.oracle.password = "<your_oracle_password>",
                    fn.oracle.dsn = "<your_oracle_dsn>"){
require(bio.portsampling)
makeHailInRpt(thePath = thePath,
              fn.oracle.username = fn.oracle.username,
              fn.oracle.password = fn.oracle.password,
              fn.oracle.dsn = fn.oracle.dsn)
  }
