#' Set the minimum and maximum display depth for an acoustic variable.
#'
#' This function changes the maximum and minimum depth displayed in the Echogram for an acoustic variable using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param acoVarName Name of acoustic variable
#' @param minDepth The minimum depth to display in metres
#' @param maxDepth The maximum depth to display in meters
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVAcoVarNameFinder}} 
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj, 'KAOS/KAOStemplate.EV')$EVFile
#' EVSetAcoVarDisplayDepth(EVFile, "38 seabed and surface excluded", 5, 250)
#'}
#'
EVSetAcoVarDisplayDepth=function (EVFile, acoVarName, minDepth, maxDepth) 
{
  if (class(minDepth) != "numeric" | class(maxDepth) != "numeric") {
    msg=paste(Sys.time(),"Error: Input numeric depths")
    warning(msg)
    return(list(msg=msg))
  }
  acoVar=EVAcoVarNameFinder(EVFile, acoVarName=acoVarName)
  msgV=acoVar$msg
  acoVar=acoVar$EVVar
  acoVarDisp=acoVar[["Properties"]][["Display"]]
  acoVarDisp[['AutoSetLimits']]=FALSE
  if(!acoVarDisp[['AutoSetLimits']]) {
    msg=paste(Sys.time(),' Display limits set to manual')
    message(msg)
    msgV=c(msgV,msg)
  } else {
    msg=paste(Sys.time(),' Failed to set display limits set to manual')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(msg=msgV))
  }
  try({
    acoVarDisp[["LowerLimit"]] <- maxDepth
  }, silent = TRUE)
  try({
    acoVarDisp[["UpperLimit"]] <- minDepth
  }, silent = TRUE)
  new_maximum <- acoVarDisp[["LowerLimit"]]
  new_minimum <- acoVarDisp[["UpperLimit"]]
  if (round(new_maximum,2) == round(maxDepth,2)) {
    msg=paste(Sys.time(),"Success: Changed the maximum depth to", 
                  maxDepth, "m")
    message(msg)
    msgV=c(msgV,msg)
  }
  else {
    msg=paste(Sys.time(),"Error: Could not change the maximum depth")
    warning(msg)
    msgV=c(msgV,msg)
  }
  if (round(new_minimum,2) == round(minDepth,2)) {
    msg=paste(Sys.time(),"Success: Changed the minimum depth to", 
                  minDepth, "m")
    message(msg)
    msgV=c(msgV,msg)
  }
  else {
    msg=paste(Sys.time(),"Error: Could not change the minimum depth")
    warning(msg)
    msgV=c(msgV,msg)
  }
  invisible(list(msg=msgV))
}
