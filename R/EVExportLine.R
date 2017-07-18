#'Export an Echoview line as a CSV file
#'
#'This function exports an Echoview line object as a CSV file.  
#'@details The Echoview line object should be associated with an Echoview acoustic variable object.  Optionally, the line can be exported for a subset of pings.
#'@param EVFile Echoview file COM object
#'@param acoVarName The acoustic variable name to which the line is associated
#'@param lineName Echoview line object name
#'@param filePath Path and filename for the resulting CSV file.
#'@param pingRange =c(-1,1) Optional ping range over which to export the Echoview line object.  Default is to export the lien for all pings
#'@export 
EVExportLineAsCSV=function(EVFile, acoVarName,lineName,filePath,pingRange=c(-1,-1))
{
  msgV=paste(Sys.time(),': Attempting to export line name',lineName,'to',filePath,'using acoustic variable',acoVarName)
  message(msgV)
  AObj=EVAcoVarNameFinder(EVFile, acoVarName)
  msgV=c(msgV,AObj$msg)
  AObj=AObj$EVVar
  
  Lobj=EVFindLineByName(EVFile=EVFile, lineName=lineName) 
  success=AObj$ExportLine(Lobj,filePath,pingRange[1],pingRange[2])
  #success=Lobj$Export(filePath)
  if(success) {msg=paste(Sys.time(), ': line,',lineName,', sucessfully exported')
              message(msg)
              msgV=c(msgV,msg)} else {
                msg=paste(Sys.time(), ': Failed to export line',lineName)
                warning(msg)
                msgV=c(msgV,msg)
              }
  invisible(list(msg=msgV))
}  


