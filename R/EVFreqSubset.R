#' Frequency subset from Wideband data.
#' 
#' This function subsets a single frequency from a wideband signal..  
#' @param EVFile An Echoview file COM object
#' @param acoVarName Acoustic variable name 
#' @param frequency A Frequency contained within the Wideband Signal (see examples).
#' @return a list object with two elements.  $success: ping subset successfully changed, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVOpenFile(EVAppObj,'~\\example1.EV')
#'EVpingSubset(EVFile=EVFile,acoVarName='Ping subset 38',pingsubsetString='2-7')
#'
#'
EVFreqSubset=function(EVFile,acoVarName,frequency)
{
  varObj<-EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=varObj$msg
  varObj<-varObj$EVVar
  msg=paste(Sys.time(),' : Changing ping subset specification in ',varObj$Name(),sep='')
  message(msg)
  msgV=c(msgV,msg)
  pSub=varObj[['Properties']][['Pingsubset']]
  message('Pre-adjust ping subset= ',pSub[['Ranges']])
  pSub[['Ranges']]<-pingsubsetString
  pSubPost=varObj[['Properties']][['Pingsubset']]$Ranges()
  msg=paste(Sys.time(),' : Post change ping subset ',pSubPost,sep='')
  message(msg)
  msgV=c(msgV,msg)
  identSubSets=identical(pingsubsetString,pSubPost)
  if(!identSubSets)
  {msg=paste(Sys.time(),' : unsuccessful ping subset change',sep= '')
  warning(msg)
  msgV=c(msgV,msg)
  }
  incExcFlag=ifelse(includePingRanges,'include','exclude')
  msg=paste(Sys.time(),' : Setting ping range to ',incExcFlag,sep='')
  message(msg)
  pSub[['RangeInclusive']]<-includePingRanges
  if(pSub[['RangeInclusive']]==includePingRanges){
    msg=paste(Sys.time(),' : Ping range inclusive successfully set to ',includePingRanges,sep='')
    message(msg)}
  else{msg=paste(Sys.time(),' : Failed to set ping range inclusive flag',sep= '')
  warning(msg)}
  msgV=c(msgV,msg)
  return(list(pingRangeSet=identSubSets,
              inclusiveFlag=pSub[['RangeInclusive']],msg=msgV))  
}  
