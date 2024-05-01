#' Get List of all variables
#' 
#' This function creates a list of all variables in the EV files
#' 
#' @param EVFile An Echoview file COM object
#' @keywords Echoview COM scripting
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' \dontrun{
#' #To be added - Needs Example data
#' #Starting Echoview
#' echoview = StartEchoview()
#' #Create a new EV File
#' EVFile <- EVCreateNew(EVAppObj=echoview, dataFiles = "rawfile.raw")$EVFile
#' Varname <- "Fileset 1: Sv pulse compressed wideband pings T1"
#' bottom <- EVBottomDetection(EVFile, EVVar=Varname, LineName="Bottom")
#' # Change the algorithm to Best bottom candidate
#' bottom <- EVBottomDetection(EVFile, EVVar=Varname, LineName="Bottom",algorithm=2)
#' }
#' @export
EVGetVariables <- function(EVFile){
  nvar <- EVFile[['Variables']]$Count() ## number of variables in EV FIles
  ## create list of variables
  vars <- vapply(seq_len(nvar), function(x) as.character(EVFile[['Variables']]$Item(x - 1)$Name()), FUN.VALUE = "")
  msgV <- if (length(vars) > 0) {
              paste0(Sys.time(), ": Success ", nvar, " variables detected")
          } else {
              paste0(Sys.time(), ": Error - No variables detected")
          }
  message(msgV)
  list(variables = vars, msg = msgV)
}
