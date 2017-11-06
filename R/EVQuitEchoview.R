#' Close an Echoview instance through COM scripting
#' 
#' This function opens an Echoview instance  
#' @param EVAppObj An Echoview Application COM object
#' @return Termination of the application and removal of the Echoview instance object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj=StartEchoview()
#'QuitEchoview(EVAppObj)
#'
#'}
QuitEchoview <- function(EVAppObj)
{
  chEVAppObj <- deparse(substitute(EVAppObj))
  if(exists(chEVAppObj)){
    EVAppObj$Quit()
    msgV=  paste(Sys.time(),': Echoview is closed...')
    message(msgV)
    rm(EVAppObj)
  }else{
    msgV=  paste(Sys.time(),': ERROR: The Echoview instance could not be found, please use a valib EVAppObj...')
    message(msgV)
  }
}  
