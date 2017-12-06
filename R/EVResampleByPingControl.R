#' Control the 'Resample by ping' Echoview virtual variable settings using COM
#' 
#' This function controls the 'Resample by ping' Echoview virtual variable settings using COM
#' 
#'@param EVFile Echoview file COM object
#'@param acoVarName name of the 'Resample by pings' acoustic variable 
#'@param NumberOfSamples number of samples to export in the depth dimension.
#'@param NumberOfPings number of pings to resample
#'@param StartRange start range, m
#'@param StopRange maximum range, m
#'@param verbose = FALSE
#'@return List of length three [1] $previousSettings [2] $currentSettings  [3] message vector
#'@export 
EVResampleByPingControl=function(EVFile,acoVarName,NumberOfSamples=NULL,NumberOfPings=NULL,StartRange=NULL,
                                 StopRange=NULL,verbose=FALSE){
  msgV=paste(Sys.time(),'Attempting to change ping resample settings in',acoVarName)
  AObj=EVAcoVarNameFinder(EVFile=EVFile, acoVarName)
  msgV=c(msgV,AObj$msg)
  AObj=AObj$EVVar
  resampObj=AObj[['Properties']][['Resample']]
  preSet=data.frame(NumberOfSamples=resampObj$NumberOfDataPoints(),
                      NumerOfPings=resampObj$NumberOfPings(),
                      StartRange=resampObj$StartRange(),
                      StopRange=resampObj$StopRange())
  
  if(!is.null(NumberOfSamples)){
  msg=paste(Sys.time(),'attempting to set number of samples')
  if(verbose) message(msg)
  msgV=c(msgV,msg)
  resampObj[['NumberOfDataPoints']]=NumberOfSamples
  if(  resampObj[['NumberOfDataPoints']]==NumberOfSamples){
  msg=paste(Sys.time(),'Success: number of samples changed to',NumberOfSamples)
  if(verbose) message(msg)
  msgV=c(msgV,msg)} else {
  msg=paste(Sys.time(),'Failed to change number of samples')
  warning(msg)
  msgV=c(msgV,msg)}
  }
  
  if(!is.null(NumberOfPings)){
    msg=paste(Sys.time(),'attempting to set number of pings')
    if(verbose) message(msg)
    msgV=c(msgV,msg)
    resampObj[['NumberOfPings']]=NumberOfPings
    if(  resampObj[['NumberOfPings']]==NumberOfPings){
      msg=paste(Sys.time(),'Success: number of samples changed to',NumberOfPings)
      if(verbose) message(msg)
      } else {
        msg=paste(Sys.time(),'Failed to change number of samples')
        warning(msg)}
    msgV=c(msgV,msg)
  }
  
  if(!is.null(StartRange)|!is.null(StopRange)){
    msg=paste(Sys.time(),'Attempting to change range selection to user defined')
    if(verbose) message(msg)
    msgV=c(msgV,msg)
    resampObj[['CustomRanges']]<-1
    if(resampObj[['CustomRanges']]==1){
      msg=paste(Sys.time(),'Success: changed range to user defined')
      if(verbose) message(msg)}else{
      msg=paste(Sys.time(),'Failed to change range to user defined')
      warning(msg)
    }
    msgV=c(msgV,msg)
    }  
  
  if(!is.null(StartRange)){
    msg=paste(Sys.time(),'attempting to set start range')
    if(verbose) message(msg)
    msgV=c(msgV,msg)
    resampObj[['StartRange']]=StartRange
    if(  resampObj[['StartRange']]==StartRange){
      msg=paste(Sys.time(),'Success: start range changed to',StartRange,'m')
      if(verbose) message(msg)
    } else {
      msg=paste(Sys.time(),'Failed to change start range')
      warning(msg)}
    msgV=c(msgV,msg)
  }
  
  if(!is.null(StopRange)){
    msg=paste(Sys.time(),'attempting to set Stop range')
    if(verbose) message(msg)
    msgV=c(msgV,msg)
    resampObj[['StopRange']]=StopRange
    if(  resampObj[['StopRange']]==StopRange){
      msg=paste(Sys.time(),'Success: Stop range changed to',StopRange,'m')
      if(verbose) message(msg)
    } else {
      msg=paste(Sys.time(),'Failed to change Stop range')
      warning(msg)}
    msgV=c(msgV,msg)
  }
  
  postSet=data.frame(NumberOfSamples=resampObj$NumberOfDataPoints(),
                    NumerOfPings=resampObj$NumberOfPings(),
                    StartRange=resampObj$StartRange(),
                    StopRange=resampObj$StopRange())
  
  invisible(list(previousSettings=preSet,currentSettings=postSet))
}



