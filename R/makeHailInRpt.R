#' @title makeHailInRpt
#' @description This function generates an excel file where the first sheet
#' lists all the groundfish hailins for the current and next day, and the
#' subsequent sheets include breakdowns of the catch.
#' @param thePath default is \code{C:\\DFO-MPO\\PORTSAMPLING}.  This is where you
#' would like your reports saved
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
#' @param usepkg default is \code{'rodbc'}. This indicates whether the connection to Oracle should
#' use \code{'rodbc'} or \code{'roracle'} to connect.  rodbc is slightly easier to setup, but
#' roracle will extract data ~ 5x faster.
#' @importFrom RODBC sqlQuery
#' @importFrom openxlsx createWorkbook
#' @importFrom openxlsx addWorksheet
#' @importFrom openxlsx writeData
#' @importFrom openxlsx saveWorkbook
#' @family Mar.portsampling
#' @author  Mike McMahon, \email{Mike.McMahon@@dfo-mpo.gc.ca}
#' @export
makeHailInRpt <- function(thePath = file.path("C:","DFO-MPO","PORTSAMPLING"),
                          fn.oracle.username = "_none_",
                          fn.oracle.password = "_none_",
                          fn.oracle.dsn = "_none_",
                          usepkg = 'rodbc', lpDays=5) {
  thePath =  path.expand(thePath)
  is.date <- function(x) inherits(x, 'POSIXct')
  fn = "PortSamplers"
  ts = format(Sys.time(), "%Y%m%d_%H%M")
  filename <- paste0(fn, "_", ts, ".xlsx")
  channel = Mar.utils::make_oracle_cxn(usepkg = usepkg, fn.oracle.username = fn.oracle.username, fn.oracle.password = fn.oracle.password, fn.oracle.dsn = fn.oracle.dsn)
  if (!inherits(channel, "OraConnection") & !inherits(channel, "RODBC")) stop("Can't connect to Oracle")
  thecmd <- Mar.utils::connectionCheck(channel)
  res <- list()

  doRpt1 <- function(){
    SQL1 = 
      "SELECT MARFISSCI.PFISP_HAIL_IN_LANDINGS.EST_LANDING_DATE_TIME,
  MARFISSCI.PFISP_HAIL_IN_LANDINGS.EST_OFFLOAD_DATE_TIME,
    MARFISSCI.COMMUNITIES.COMMUNITY_NAME
    || ' ('
    || MARFISSCI.WHARVES.WHARF_NAME
    || ')' PORT_WHARF,
    MARFISSCI.VESSELS.VESSEL_NAME,
    MARFISSCI.GEARS.DESC_ENG AS GEAR,
    MARFISSCI.PFISP_HAIL_IN_CALLS.VR_NUMBER,
    MARFISSCI.PFISP_HAIL_OUT_LIC_DOCS.LICENCE_ID,
    MARFISSCI.HAIL_IN_TYPES.DESC_ENG OFFLOAD,
    SUM(ROUND((
    CASE MARFISSCI.PFISP_HAIL_IN_ONBOARD.UNIT_OF_MEASURE_ID
    WHEN 10
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2.20462
    WHEN 20
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    WHEN 30
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2204.62
    WHEN 40
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2000
    ELSE MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    END), 1)) TOT_EST_ONBOARD_WEIGHT_LBS,
    ALLSPEC.SPECIES,
    REGEXP_REPLACE(ALLSPEC.NAFO_ZONES, '([^,]+)(,[ ]*\\1)+', '\\1') NAFO_ZONES,
    MARFISSCI.DMP_COMPANIES.NAME DMP_NAME,
    MARFISSCI.DOCKSIDE_OBSERVERS.FIRSTNAME
    || ' '
    || MARFISSCI.DOCKSIDE_OBSERVERS.SURNAME DS_OBS,  
    MARFISSCI.PFISP_HAIL_IN_CALLS.OBSERVER_FLAG,
    MARFISSCI.PFISP_HAIL_IN_CALLS.CONF_NUMBER,
    MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_LANDING_ID
    FROM
    MARFISSCI.PFISP_HAIL_IN_LANDINGS,
    MARFISSCI.PFISP_HAIL_IN_CALLS,
    MARFISSCI.WHARVES,
    MARFISSCI.COMMUNITIES,
    MARFISSCI.HAIL_IN_TYPES,
    MARFISSCI.DMP_COMPANIES,
    MARFISSCI.PFISP_HAIL_IN_ONBOARD,
    MARFISSCI.PFISP_HAIL_OUTS,
    MARFISSCI.VESSELS,
    MARFISSCI.SPECIES,
    MARFISSCI.PFISP_HAIL_OUT_LIC_DOCS,
    MARFISSCI.LICENCES,
    MARFISSCI.PFISP_HAIL_IN_LAND_OBSRVRS,
    MARFISSCI.DOCKSIDE_OBSERVERS,
    (SELECT HAIL_IN_LANDING_ID,
    ListAgg(DESC_ENG, ', ') Within GROUP (
    ORDER BY RANCOR) SPECIES,
    ListAgg(AREA, ', ') Within GROUP (
    ORDER BY RANCOR) NAFO_ZONES
    FROM
    (SELECT ANAL.HAIL_IN_LANDING_ID,
    REGEXP_REPLACE(SP.DESC_ENG, ',', '-') DESC_ENG,
    NAFO.AREA,
    ANAL.RANCOR
    FROM
    (SELECT MARFISSCI.PFISP_HAIL_IN_ONBOARD.HAIL_IN_LANDING_ID,
    MARFISSCI.PFISP_HAIL_IN_ONBOARD.SSF_SPECIES_CODE,
    (
    CASE MARFISSCI.PFISP_HAIL_IN_ONBOARD.UNIT_OF_MEASURE_ID
    WHEN 10
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2.20462
    WHEN 20
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    WHEN 30
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2204.62
    WHEN 40
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2000
    ELSE MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    END) EST_ONBOARD_WEIGHT_LBS,
    MARFISSCI.PFISP_HAIL_IN_ONBOARD.NAFO_UNIT_AREA_ID,
    rank() Over (Partition BY MARFISSCI.PFISP_HAIL_IN_ONBOARD.HAIL_IN_LANDING_ID Order By (
    CASE MARFISSCI.PFISP_HAIL_IN_ONBOARD.UNIT_OF_MEASURE_ID
    WHEN 10
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2.20462
    WHEN 20
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    WHEN 30
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2204.62
    WHEN 40
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2000
    ELSE MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    END) DESC) RANCOR
    FROM MARFISSCI.PFISP_HAIL_IN_ONBOARD
    ) ANAL,
    MARFISSCI.SPECIES SP,
    MARFISSCI.NAFO_UNIT_AREAS NAFO
    WHERE ANAL.SSF_SPECIES_CODE = SP.SPECIES_CODE(+)
    AND ANAL.NAFO_UNIT_AREA_ID  = NAFO.AREA_ID(+)
    )
    GROUP BY HAIL_IN_LANDING_ID
    ) ALLSPEC,
    MARFISSCI.PRO_SPC_INFO,
    MARFISSCI.GEARS
    WHERE MARFISSCI.PFISP_HAIL_IN_CALLS.HAIL_IN_CALL_ID           = MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_CALL_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LANDINGS.WHARF_ID                 = MARFISSCI.WHARVES.WHARF_ID(+)
    AND MARFISSCI.WHARVES.COMMUNITY_CODE                          = MARFISSCI.COMMUNITIES.COMMUNITY_CODE(+)
    AND MARFISSCI.PFISP_HAIL_IN_CALLS.HAIL_IN_TYPE_ID             = MARFISSCI.HAIL_IN_TYPES.HAIL_IN_TYPE_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LANDINGS.TRIP_DMP_COMPANY_ID      = MARFISSCI.DMP_COMPANIES.DMP_COMPANY_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_LANDING_ID       = MARFISSCI.PFISP_HAIL_IN_ONBOARD.HAIL_IN_LANDING_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_CALLS.HAIL_OUT_ID                 = MARFISSCI.PFISP_HAIL_OUTS.HAIL_OUT_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_CALLS.VR_NUMBER                   = MARFISSCI.VESSELS.VR_NUMBER(+)
    AND MARFISSCI.PFISP_HAIL_OUTS.HAIL_OUT_ID                     = MARFISSCI.PFISP_HAIL_OUT_LIC_DOCS.HAIL_OUT_ID(+)
    AND MARFISSCI.PFISP_HAIL_OUT_LIC_DOCS.LICENCE_ID              = MARFISSCI.LICENCES.LICENCE_ID(+)
    AND MARFISSCI.LICENCES.SPECIES_CODE                           = MARFISSCI.SPECIES.SPECIES_CODE(+)
    AND MARFISSCI.PFISP_HAIL_IN_CALLS.TRIP_DMP_COMPANY_ID         = MARFISSCI.PFISP_HAIL_IN_LANDINGS.TRIP_DMP_COMPANY_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LAND_OBSRVRS.DOCKSIDE_OBSERVER_ID = MARFISSCI.DOCKSIDE_OBSERVERS.DOCKSIDE_OBSERVER_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LAND_OBSRVRS.TRIP_DMP_COMPANY_ID  = MARFISSCI.DOCKSIDE_OBSERVERS.DMP_COMPANY_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_LANDING_ID       = MARFISSCI.PFISP_HAIL_IN_LAND_OBSRVRS.HAIL_IN_LANDING_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_LANDING_ID       = ALLSPEC.HAIL_IN_LANDING_ID(+)
    AND MARFISSCI.PFISP_HAIL_IN_CALLS.TRIP_ID                     = MARFISSCI.PRO_SPC_INFO.TRIP_ID(+)
    AND MARFISSCI.PRO_SPC_INFO.GEAR_CODE                          = MARFISSCI.GEARS.GEAR_CODE(+)
    AND MARFISSCI.PFISP_HAIL_IN_LANDINGS.EST_LANDING_DATE_TIME BETWEEN TRUNC(SysDate - 1) AND TRUNC(SysDate + 2)
    AND MARFISSCI.SPECIES.SPECIES_CATEGORY_ID = 1
    GROUP BY
    MARFISSCI.PFISP_HAIL_IN_LANDINGS.EST_LANDING_DATE_TIME,
    MARFISSCI.PFISP_HAIL_IN_LANDINGS.EST_OFFLOAD_DATE_TIME,
    MARFISSCI.VESSELS.VESSEL_NAME,
    MARFISSCI.PFISP_HAIL_IN_CALLS.VR_NUMBER,
    MARFISSCI.PFISP_HAIL_OUT_LIC_DOCS.LICENCE_ID,
    MARFISSCI.COMMUNITIES.COMMUNITY_NAME
    || ' ('
    || MARFISSCI.WHARVES.WHARF_NAME
    || ')',
    MARFISSCI.HAIL_IN_TYPES.DESC_ENG,
    ALLSPEC.SPECIES,
    MARFISSCI.DMP_COMPANIES.NAME,
    MARFISSCI.PFISP_HAIL_IN_CALLS.CONF_NUMBER,  
    MARFISSCI.PFISP_HAIL_IN_CALLS.OBSERVER_FLAG,
    MARFISSCI.PFISP_HAIL_IN_CALLS.CONF_ISSUED_DATE_TIME
    || '('
    || MARFISSCI.PFISP_HAIL_IN_CALLS.CONF_ISSUED_USER_ID
    || ')',
    MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_CALL_ID,
    MARFISSCI.PFISP_HAIL_IN_CALLS.HAIL_OUT_ID,
    MARFISSCI.PFISP_HAIL_IN_CALLS.TRIP_ID,
    MARFISSCI.PFISP_HAIL_IN_LANDINGS.TRIP_DMP_COMPANY_ID,
    MARFISSCI.SPECIES.SPECIES_CATEGORY_ID,
    MARFISSCI.LICENCES.LICENCE_ID,
    MARFISSCI.DOCKSIDE_OBSERVERS.FIRSTNAME,
    MARFISSCI.DOCKSIDE_OBSERVERS.SURNAME,
    ALLSPEC.NAFO_ZONES,
    MARFISSCI.PFISP_HAIL_OUT_LIC_DOCS.GEAR_CODE,
    MARFISSCI.GEARS.DESC_ENG,
    MARFISSCI.PFISP_HAIL_IN_LANDINGS.HAIL_IN_LANDING_ID
    ORDER BY MARFISSCI.PFISP_HAIL_IN_LANDINGS.EST_LANDING_DATE_TIME DESC,
    MARFISSCI.PFISP_HAIL_IN_CALLS.VR_NUMBER"
    
    data = thecmd(channel, SQL1)
    if (nrow(data) == 0) stop("No data returned")
    cat("\nReceived data")
    data[,!sapply(data, is.date)][is.na(data[,!sapply(data, is.date)])] <- 0 
    data[, sapply(data, is.date)][is.na(data[, sapply(data, is.date)])] <- as.Date('9999/01/01')
    
    return(data)
  }
  
  doRpt2 <- function(HILID = NULL){
    # thisSQLDET = gsub("&HILID&", HILID$HAIL_IN_LANDING_ID[x], SQLDET)
    SQLDET = paste0(
      "SELECT
    SPECIES.DESC_ENG SPECIES,
    PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT,
    UNIT_OF_MEASURES.DESC_ENG UNITS,
    PFISP_HAIL_IN_ONBOARD.SSF_LANDED_FORM_CODE FORM_CODE,
    NAFO_UNIT_AREAS.AREA NAFO_AREA,
    PFISP_HAIL_IN_CALLS.VR_NUMBER,
    PFISP_HAIL_IN_ONBOARD.HAIL_IN_LANDING_ID
    FROM MARFISSCI.PFISP_HAIL_IN_ONBOARD
    LEFT JOIN MARFISSCI.SPECIES
    ON SPECIES.SPECIES_CODE = PFISP_HAIL_IN_ONBOARD.SSF_SPECIES_CODE
    LEFT JOIN MARFISSCI.UNIT_OF_MEASURES
    ON UNIT_OF_MEASURES.UNIT_OF_MEASURE_ID = PFISP_HAIL_IN_ONBOARD.UNIT_OF_MEASURE_ID
    LEFT JOIN MARFISSCI.NAFO_UNIT_AREAS
    ON PFISP_HAIL_IN_ONBOARD.NAFO_UNIT_AREA_ID = NAFO_UNIT_AREAS.AREA_ID
    RIGHT JOIN MARFISSCI.PFISP_HAIL_IN_LANDINGS
    ON PFISP_HAIL_IN_ONBOARD.HAIL_IN_LANDING_ID = PFISP_HAIL_IN_LANDINGS.HAIL_IN_LANDING_ID
    RIGHT JOIN MARFISSCI.PFISP_HAIL_IN_CALLS
    ON PFISP_HAIL_IN_LANDINGS.HAIL_IN_CALL_ID      = PFISP_HAIL_IN_CALLS.HAIL_IN_CALL_ID
    WHERE PFISP_HAIL_IN_ONBOARD.HAIL_IN_LANDING_ID = ",HILID,"
    ORDER BY
    (
    CASE MARFISSCI.PFISP_HAIL_IN_ONBOARD.UNIT_OF_MEASURE_ID
    WHEN 10
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2.20462
    WHEN 20
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    WHEN 30
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2204.62
    WHEN 40
    THEN MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT * 2000
    ELSE MARFISSCI.PFISP_HAIL_IN_ONBOARD.EST_ONBOARD_WEIGHT
    END) DESC"
    )
    data = thecmd(channel, SQLDET)
    return(data)
  }
  
  doLargePelagics <- function(lpDays=5){
    SQLTuna = paste0("SELECT DISTINCT
    S.DESC_ENG SPECIES,
    A.EST_ONBOARD_WEIGHT,
    A.OFFLOAD_WEIGHT,
    N.AREA NAFO,
    A.VR_NUMBER,
    V.VESSEL_NAME, 
    A.SPC_TAG_ID,
    C.COMMUNITY_NAME,
    W.WHARF_NAME,
    B.EST_OFFLOAD_DATE_TIME,
    B.EST_LANDING_DATE_TIME,
    DMP.NAME,
    DO.FIRSTNAME || ' ' || DO.SURNAME DS_OBS,
    HIC.CONF_NUMBER,
    HIC.OBSERVER_FLAG   
FROM
    MARFISSCI.HAIL_IN_ONBOARD    A,
    MARFISSCI.HAIL_IN_LANDINGS   B,
    MARFISSCI.COMMUNITIES C,
    MARFISSCI.WHARVES W,
    MARFISSCI.VESSELS V,
    MARFISSCI.SPECIES S,
    MARFISSCI.NAFO_UNIT_AREAS N,
    MARFISSCI.HAIL_IN_CALLS HIC,
    MARFISSCI.HAIL_IN_LAND_OBSRVRS HILO,
    MARFISSCI.DOCKSIDE_OBSERVERS DO,
    MARFISSCI.DMP_COMPANIES DMP
WHERE
    A.HAIL_IN_LANDING_ID = B.HAIL_IN_LANDING_ID
    AND B.COMMUNITY_CODE = C.COMMUNITY_CODE
    AND B.WHARF_ID = W.WHARF_ID
    AND A.VR_NUMBER = V.VR_NUMBER
    AND  A.SSF_SPECIES_CODE = S.SPECIES_CODE 
    AND A.NAFO_UNIT_AREA_ID = N.AREA_ID
    AND B.HAIL_IN_CALL_ID = HIC.HAIL_IN_CALL_ID
    AND A.HAIL_IN_LANDING_ID = HILO.HAIL_IN_LANDING_ID(+)
    AND HILO.DOCKSIDE_OBSERVER_ID = DO.DOCKSIDE_OBSERVER_ID(+)
    AND HILO.TRIP_DMP_COMPANY_ID = DMP.DMP_COMPANY_ID(+)
    AND (B.EST_LANDING_DATE_TIME >= sysdate-",lpDays," OR B.EST_OFFLOAD_DATE_TIME >= sysdate-",lpDays,") 
    AND SSF_SPECIES_CODE IN (
   251, --swordfish
   252, --albacore 
   253, --bigeye
   254, --bluefin
   255, --skipjack
   256, --yellowfin
   257, --tuna,restricted
   259  --tuna, unspecified
    )
    ORDER BY GREATEST(EST_OFFLOAD_DATE_TIME,EST_LANDING_DATE_TIME) DESC")
    data = thecmd(channel, SQLTuna)
    return(data)
  }
  
  makeHILID <- function(data = NULL){
    data$tmpID <- seq.int(nrow(data))
    HILID = data[data$TOT_EST_ONBOARD_WEIGHT_LBS > 0, c("tmpID", "VESSEL_NAME", "EST_LANDING_DATE_TIME","HAIL_IN_LANDING_ID")]
    HILID$tmpEST_LANDING_DATE_TIME = gsub(':|-','',HILID$EST_LANDING_DATE_TIME)
    HILID$tmpEST_LANDING_DATE_TIME = substr(gsub(' ','_',HILID$tmpEST_LANDING_DATE_TIME),7,13)
    HILID = HILID[with(HILID, order(HILID$VESSEL_NAME,HILID$tmpEST_LANDING_DATE_TIME)),]
    HILID$tmpVESS_INFO = paste0(substr(HILID$VESSEL_NAME,1,23),"_",HILID$tmpEST_LANDING_DATE_TIME)
    HILID$tmpCNT = sequence(rle(as.character(HILID$tmpVESS_INFO))$lengths)
    #vessels that show up more than once should be identified
    if (max(HILID$tmpCNT)>1){
      HILID[HILID$tmpCNT > 1, ]$tmpVESS_INFO <- paste0(HILID[HILID$tmpCNT > 1, ]$tmpVESS_INFO, "_", HILID[HILID$tmpCNT > 1, ]$tmpCNT)
    }
    HILID = HILID[with(HILID, order(HILID$tmpID)),]
    HILID$tmpID <- NULL
    HILID$tmpEST_LANDING_DATE_TIME <- NULL
    HILID$tmpCNT <- NULL
    return(HILID)
  }

  data <- doRpt1()
  if(nrow(data)==0){
    data[1,] <- NA
  }
  lpelagics <- doLargePelagics()
  if(nrow(lpelagics)==0){
    lpelagics[1,] <- NA
  }
  HILID <- makeHILID(data=data)
  if(nrow(HILID)==0){
    HILID[1,] <- NA
  }
  
  #remove illegal stuff from the field that will become excel sheet names
  HILID$tmpVESS_INFO <-gsub(pattern = "_+", "_", gsub("[[:punct:]]|[[:blank:]]", "_", HILID$tmpVESS_INFO, fixed = F))
  # dir.create(thePath, showWarnings = FALSE)
  wb <- openxlsx::createWorkbook()
  if(!is.null(data)) {
    openxlsx::addWorksheet(wb, "MASTER")
    openxlsx::writeData(wb, sheet = "MASTER", data, rowNames = F)
  }
  if(!is.null(lpelagics)) {
    openxlsx::addWorksheet(wb, "LARGEPELAGICS")
    openxlsx::writeData(wb, sheet = "LARGEPELAGICS", lpelagics, rowNames = F)
  }

  #write data to sheet
  cat("\nGetting details")
  for (x in 1:nrow(HILID)) {
    datadet <- doRpt2(HILID = HILID$HAIL_IN_LANDING_ID[x])
    if (nrow(datadet)>0){
      
      openxlsx::addWorksheet(wb, as.character(HILID$tmpVESS_INFO[x]))
      openxlsx::writeData(wb, sheet = as.character(HILID$tmpVESS_INFO[x]), datadet, rowNames = F)
    }
  }
  openxlsx::saveWorkbook(wb, file.path(thePath,filename), overwrite = T)
  cat(paste0("\nFile written to ",file.path(thePath,filename),"\n\n"))
}
