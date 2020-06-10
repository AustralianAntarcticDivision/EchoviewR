#' Get List of all variables
#' 
#' This function creates a list of all variables in the EV files
#' 
#' @param EVFile An Echoview file COM object
#' @keywords Echoview COM scripting
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' #To be added - Needs Example data
#' #Starting Echoview
#' echoview = StartEchoview()
#' #Create a new EV File
#' EVFile <- EVCreateNew(EVAppObj=echoview, dataFiles = "rawfile.raw")$EVFile
#' Varname <- "Fileset 1: Sv pulse compressed wideband pings T1"
#' bottom <- EVBottomDetection(EVFile, EVVar=Varname, LineName="Bottom")
#' # Change the algorithm to Best bottom candidate
#' bottom <- EVBottomDetection(EVFile, EVVar=Varname, LineName="Bottom",algorithm=2)
#' @export

EVGetVariables <- function(EVFile){
  #number of variables in EV FIles
  nvar = EVFile$EVFile[['Variables']]$Count()
  #create list of variables
  vars = apply(matrix(1:nvar),1,
        FUN=function(x){EVFile$EVFile[['Variables']]$Item(x)$Name()})
  if(length(vars) > 0){
    msgV = paste0(Sys.time(),": Success ", nvar, " variables detected")
  }else{
    msgV = paste0(Sys.time(),": Error - No variables detected")
  }
  message(msgV)
        
  return(list(variables=vars,msg=msgV))
}
  