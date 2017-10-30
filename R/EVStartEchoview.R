#' Start an Echoview instance through COM scripting
#' 
#' This function opens an Echoview instance  
#' @return an object of class COMDispatch
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj=StartEchoview()
#'
#'}
StartEchoview <- function()
{
  msgV=  paste(Sys.time(),': Echoview is running...')
  message(msgV)
  return(COMCreate('EchoviewCom.EvApplication'))
}  
