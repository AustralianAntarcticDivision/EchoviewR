#' Find an acoustic variable by name
#' 
#' This function finds an acoustic variable in an Echoview file by name and 
#' returns the variable pointer.
#' @param EVFile An Echoview file COM object
#' @param acoVarName The name of an acoustic variable in the Echoview file
#' @return a list object with two elements. $EVVar: An Echoview acoustic variable object, and $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, "120 7x7 convolution")$EVVar
#' }
EVAcoVarNameFinder <- function (EVFile, acoVarName) {
  
  obj <- EVFile[["Variables"]]$FindByName(acoVarName)
  if (is.null(obj)) {
    obj <- EVFile[["Variables"]]$FindByShortName(acoVarName)
  }
  if (is.null(obj)) {
    msg <- paste(Sys.time(), ' : Variable not found ', acoVarName, sep = '')
    warning(msg)
    obj <- NULL
  } else {
    msg <- paste(Sys.time(), ' : Variable found ', acoVarName, sep = '') 
    message(msg)
  }
  return(list(EVVar = obj, msg = msg))  
}


