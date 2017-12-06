
#' Sets maximum data threshold for a variable object
#' 
#' This function sets the maxmimum data threshold
#' @param EVFile An Echoview file COM object
#' @param acoVarName (character) Echoview acoustic variable name
#' @param thres The new maximum threshold to be set (dB)
#' @return a list object with two elements. $thresholdSettings: The new threshold settings; $msg vector of messages
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVminThresholdSet}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, "38 seabed and surface excluded")$EVVar
#' EVmaxThresholdSet(varObj, -30)
#' }
EVmaxThresholdSet <- function (EVFile,acoVarName, thres,colRng=36) {
acousticVar=EVAcoVarNameFinder(EVFile, acoVarName = acoVarName)
  msgV=paste(Sys.time(),' : Attempting to change maximum data threshold in ',acoVarName)
  message(msgV)
  msgV=c(msgV,acousticVar$msg)
  varObj=acousticVar$EVVar

  varDat <- varObj[["Properties"]][["Data"]]
  preThresApplyFlag <- varDat$ApplyMaximumThreshold()
  varDat[['ApplyMaximumThreshold']] <- TRUE
  postThresApplyFlag <- varDat$ApplyMaximumThreshold()
  
  varDat[['LockSvMaximum']] <-FALSE
  if (postThresApplyFlag) {
    msg <- paste(Sys.time(),' : Apply maximum threshold flag set to TRUE in ',
                 acoVarName, sep = '')
    	msgV=c(msgV,msg)
	message(msg)
  } else {
    msg <- paste(Sys.time(),' : Failed to set maximum threshold flag in ',
                 acoVarName, sep = '')
    warning(msg)
	msgV=c(msgV,msg)
    return(list(msg=msgV,thresholdSettings=NA))
  }
  
  #set threshold value
  preMaxThresVal <- varDat$MaximumThreshold()
  varDat[['MaximumThreshold']] <- thres
  postMaxThresVal <- varDat$MaximumThreshold()
  if (postMaxThresVal == thres) {
    msg2 <- paste(Sys.time(), ' : Maximum threshold successfully set to ', thres, 
                  ' in ', acoVarName, sep = '')
    message(msg2)
    msgV <- c(msg, msg2)
  } else {
    msg2 <- paste(Sys.time(), ' : Failed to set minimum threshold in ', 
                  acoVarName, sep = '')
    warning(msg2)
    msgV <- c(msg, msg2)
	return(list(msg=msgV,thresholdSettings=NA))
  }
  
  #now try with display threshold
  msg=paste(Sys.time(),' : attempting to change maximum display threshold using display colour range')
  msgV=c(msgV,msg)
  message(msg)
  varDisp <- varObj[["Properties"]][['Display']]
  minDisp=varDisp[['ColorMinimum']]
  colDiff=-(minDisp-thres)
  
  if(!is.na(colRng)) 
  {colDiff[colDiff>colRng]=colRng
  if(colDiff>colRng){
  msg=paste(Sys.time(),' : colour range exceeded so not rescaling.  If you want to rescale anything, change the colRng ARG to colRng=NA')
  message(msg)
  msgV=c(msgV,msg)}
    }
  
  varDisp[['ColorRange']]=colDiff
  
  maxCol=varDisp[['ColorRange']] + varDisp[['ColorMinimum']]
  
  if(maxCol == thres){
    msg=paste(Sys.time(), ' : success changed maximum display colour by adjusting display range')
    msgV=c(msgV,msg)
    message(msg)
  } else
  {
    msg=paste(Sys.time(), ' : Failed to change maximum display colour.')
    msgV=c(msgV,msg)
    message(msg)
  }  
  
  return(list(thresholdSettings = c(preThresApplyFlag = preThresApplyFlag, 
                                    preMaxThresVal = preMaxThresVal,
                                    postThresApplyFlag = postThresApplyFlag,
                                    postMaxThresVal = postMaxThresVal), msg = msgV))
}  

