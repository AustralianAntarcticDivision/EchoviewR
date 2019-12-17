#' Get the current version of Echoview
#' 
#' This function returns the version of Echoview that was started through EchoviewR
#' 
#' @param EVApp An Echoview Application object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' #Create an Echoview application object and start Echoview
#' EVApp <- StartEchoview()
#' #get the current version of the Echoview App started through EchoviewR
#' EVVersion(EVApp)
#' #Quit the application again
#' QuitEchoview(EVApp)
#' 


EVVersion <- function(EVApp=NULL){
  if(!is.null(EVApp)){
    #Get the version
    ver <- EVApp[['Version']]
    message(paste0(Sys.time(),": The version of Echoview currently running through EchoviewR is ", ver))
  }else{
    message(paste0(Sys.time(),": Error - Please provide a valid Echoview application object"))
  }
  
}

