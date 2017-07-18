#' Change the top and bottom exclusion lines in an Echoview acoustic variable
#' 
#' This function changes the top and bottom exclusion lines in an Echoview acoustic variable using COM scripting.  
#' @param EVFile An Echoview file COM object
#' @param acoVarName Acoustic variable name.
#' @param newAboveExclusionLine Name of the top exclusion line to be used. Default is NULL in which case the top exclusion line is unchanged.
#' @param newBelowExclusionLine Name of the bottom exclusion line to be used. Default is NULL in which case the bottom exclusion line is unchanged.
#' @return a list object with two elements.  $analysisProperties: The analysis properties COM object specified in \code{acoVarName}, and $msg: message vector for the processing log. 
#' @keywords Echoview COM scripting
#' @author Martin Cox
#' @seealso \link{EVAcoVarNameFinder}
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
EVsetExclusionLines<-function(EVFile,acoVarName,newAboveExclusionLine=NULL,newBelowExclusionLine=NULL)
{
  #Find the acoustic variable:
  finder=EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=finder$msg #we will keep track of messages produced by the function here
  acoVar=finder$EVVar #this is the COM address of the acoustic virtual variable of interest
  
  #get the analysis properties object of the EV acoustic variable of interest:
  anaObj=acoVar[['Properties']][['Analysis']]
  #check that there are analysis properties available for the selected acoustic variable:
  if(class(anaObj)[1]!='COMIDispatch'){
    #return a warning message if there acoustic variable does not have analysis properties:
    msg=paste(Sys.time(),'No analysis properties found in acoustic variable', acoVar$Name())
    #the following line is a vector of messages that are displyed as we progress through the function-
    #= and is returned at the end of the function, or if the function ends early due to an error.
    msgV=c(msgV,msg)
    warning(msg)
    return(list(msg=msgV))
  } #end of analysis properties check.
  
  #process the surface exclusion line:
  if(!is.null(newAboveExclusionLine))
  {
    msg=paste(Sys.time(),'Changing surface exclusion line name from',
              anaObj$ExcludeAboveLine() ,'to',newAboveExclusionLine)
    message(msg);msgV=c(msgV,msg)
    #change line name:
    anaObj[['ExcludeAboveLine']]<-newAboveExclusionLine
    #check line change was successful:
    if(anaObj$ExcludeAboveLine()==newAboveExclusionLine){
      msg=paste(Sys.time(),' Surface exclusion line name changed to ',
                anaObj$ExcludeAboveLine())
      message(msg);msgV=c(msgV,msg)} else {
        msg=paste(Sys.time(),' Failed to change surface exclusion line')
        warning(msg);msgV=c(msgV,msg)
        return(list(msg=msgV))}
  } #end of processing surface exclusion line
  
  #process the bottom exclusion line:
  if(!is.null(newBelowExclusionLine))
  {
    msg=paste(Sys.time(),'Changing bottom exclusion line name from',
              anaObj$ExcludeBelowLine() ,'to',newBelowExclusionLine)
    message(msg);msgV=c(msgV,msg)
    #change line name:
    anaObj[['ExcludeBelowLine']]<-newBelowExclusionLine
    #check line change was successful:
    if(anaObj$ExcludeBelowLine()==newBelowExclusionLine){
      msg=paste(Sys.time(),' Below exclusion line name changed to ',anaObj$ExcludeBelowLine())
      message(msg);msgV=c(msgV,msg)} else {
        msg=paste(Sys.time(),' Failed to change below exclusion line')
        warning(msg);msgV=c(msgV,msg)
        return(list(msg=msgV))}
  } #end of processing below exclusion line
  return(list(analysisProperties=anaObj,msg=msgV))
} #end of  EVsetExclusionLines function
