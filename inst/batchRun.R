args=(commandArgs(TRUE))
require(bio.portsampling)
makeHailInRpt(thePath = "C:/DFO-MPO/PORTSAMPLING",
              fn.oracle.username = args[1],
              fn.oracle.password = args[2],
              fn.oracle.dsn = args[3])
