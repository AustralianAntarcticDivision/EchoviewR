#' Export echo  integration from Echoview for all regions in an Echoview acoustic variable
#' 
#' This function integrates all acoustic data within all regions for an Echoview acoustic variable.  The data can either be exported on a grid or as one integration per region.
#' @param EVFile ("COMIDispatch) An Echoview file COM object
#' @param acoVarName (string) Echoview acoustic variable name
#' @param grid = FALSE (boolean) Integrate all data within a region to the region extents or when TRUE integrate on to a grid within the Echoview acoustic variable
#' @param exportFn CSV export file path and file name.
#' @return list $msg messages resulting from function call
#' @export
#' @seealso \code{\link{EVChangeVariableGrid}} 
EVExportIntegrationAllRegions=function(EVFile,acoVarName,grid=FALSE,exportFn)
{
  eVVar=EVAcoVarNameFinder(EVFile = EVFile, acoVarName=acoVarName)
  obj=eVVar$EVVar;msgV=eVVar$msg
  if(!grid){
    msg=paste(Sys.time(),' :Integrating all data within all regions for acoustic variable',acoVarName)
    message(msg)
    msgV=c(msgV,msg)
    success=obj$ExportIntegrationByRegionsAll(exportFn)
  }
  if(grid){
    msg=paste(Sys.time(),': Integrating data on to a grid within all regions for acoustic variable',acoVarName)
    message(msg)
    msgV=c(msgV,msg)
    success=obj$ExportIntegrationByRegionsByCellsAll(exportFn)
  }
  
  
  if(success) {
    msg=paste(Sys.time(),': Integrated Sv data from',acoVarName)
    message(msg)
    msgV=c(msgV,msg)
  } else {
    if(grid){
      msg=paste(Sys.time(),': Echoview failed to export regions.  Perhaps check the integration grid settings in ',acoVarName)} else{
        msg=paste(Sys.time(),': Echoview failed to export regions.')}
  
    warning(msg)
    msgV=c(msgV,msg)
    return(list(msg=msgV))
  }
  if(!file.exists(exportFn)){
    msg=paste(Sys.time(),': Exported file not found. Check file write permissions')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(msg=msgV))
  }
  invisible(list(msg=msgV))
}
