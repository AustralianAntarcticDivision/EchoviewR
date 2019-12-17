#' Start Echoview over COM 
#' 
#' This function opens an Echoview instance through COM scripting
#' @return an Echoview Application object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' EVAppObj <- StartEchoview()

StartEchoview <- function () {
  message(paste0(Sys.time(),": Opening Echoview"))
  EVApp=COMCreate('EchoviewCom.EvApplication') 
  message(paste0(Sys.time(),": Opened Echoview version ",EVApp$Version()))
  scriptLic=EVApp$IsLicensed()
  if(!scriptLic) warning(paste0(Sys.time(),": Echoview is NOT licensed for scripting"))
  return(EVApp)
}

#' Quit Echoview over COM 
#' 
#' This function closes an Echoview instance through COM scripting
#' @param EVAppObj Echoview Application Object
#' @return an Echoview Application object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' EVAppObj <- StartEchoview()
QuitEchoview <- function (EVAppObj) {
  message(paste0(Sys.time(),": Closing Echoview"))
  EVAppObj$Quit()
}
