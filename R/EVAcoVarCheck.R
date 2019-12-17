#' Check an Echoview acoustic variable
#' 
#' This function checks if an Echoview acoustic variable is useable and returns the number of pings in the variable.  
#' 
#' @param EVFile (COM object) of the open Echoview File
#' @param acoVarName (character) Acoustic variable name
#' @return (list) $useable (boolean) $measurementCount (integer) number of pings in the acoustic variable NA if unuseable $msg (character) vector of messages
#' @export 
EVAcoVarCheck=function(EVFile,acoVarName){
  varRes=EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=varRes$msg
  obj=varRes$EVVar
  if(is.null(obj)){
    msg=paste(Sys.time(),' : ',acoVarName,' not found. Quitting')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(useable=FALSE,measurementCount=NA,msg=msgV))
  }
  #check variable type:
  if(obj$VariableType()!=1) {
    msg= paste(Sys.time(),' : ',acoVarName,' is not an acoustic variable. Quitting')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(useable=FALSE,measurementCount=NA,msg=msgV))
  }

  msg=paste(Sys.time(),': Checking acoustic variable called ',acoVarName)
  message(msg)
  msgV=c(msgV,msg)
  useableFLAG=obj$Useable()
  msg=paste(Sys.time(),':',acoVarName,ifelse(useableFLAG,'is', 'is  not'),'useable')
  message(msg)
  msgV=c(msgV,msg)
  if(useableFLAG)
  {
    measurementCount=obj$MeasurementCount()
    msg=paste(Sys.time(),':',acoVarName,'has',measurementCount,'pings')
    message(msg)
    msgV=c(msgV,msg)} else {measurementCount=NA}
  return(list(useable=useableFLAG,measurementCount=measurementCount,msg=msgV))
}
