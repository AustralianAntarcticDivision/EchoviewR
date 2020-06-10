#' Adjust the Echoview virtual variable 'Background noise removal' using COM
#'
#' This function adjusts the settings of a Background Noise Removal virtual variable using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param acoVarName (character) Echoview acoustic variable name
#' @param timeDistanceGridType (numeric) specifying the along track grid type. 0 = no grid, 1 = time (minutes), 2 = GPS distance (NMi), 3 = Vessel Log Distance (NMi), 4 = Pings, 5 = GPS distance (m), 6 = Vessel Log Distance (m). 
#' @param depthGridType (numeric) 0 = no grid, 1 = depth grid, 2 = use reference line.
#' @param timeDistanceGridDistance (numeric) vertical grid line spacing. Not needed if verticalType = 0. 
#' @param depthGridDistance (numeric) horizontal grid line spacing. Ignored if horizontalType = 0. 
#' @param EVLineName (character) an Echoview line object name. Not needed if depthGridType != 2.
#' @param showenum = FALSE (boolean) if TRUE Echoview grid enum values are returned to the workspace.  All other arguments are ignored.
#' @return a list object with three elements.  $predataRangeSettings: a vector of pre-function call settings; $postDataRangeSettings a vector of post-function call settings, and $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @details \code{depthGridType} is specified in the Echoview COM enum values called 'EDepthRangeGridMode'. \code{TimeDistanceGridMode} is specified in the Echoview COM enum values called ETimeDistanceGridMode
#' @export
#' @author Martin Cox
#' @note This function was completely rewritten Sept. 2017 which resulted in changes to function argument names.
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVAcoVarNameFinder}} \code{\link{EVFindLineByName}}
#' @examples
#' \dontrun{
#' EVAppObj <- start
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' EVChangeVariableGrid <- function (EVFile, 
#' acoVarName="38 seabed and surface excluded", 
#' timeDistanceGridType=5, depthGridType=2, 
#' timeDistanceGridDistance = 100, depthGridDistance = 10, 
#' EVLineName = "Fixed depth 250 m")
#' #remove grid:
#' EVChangeVariableGrid <- function (EVFile, 
#' acoVarName="38 seabed and surface excluded", 
#' timeDistanceGridType=0, depthGridType=0)
#' #only show Echoview grid enum values
#' EVChangeVariableGrid(showenum = TRUE)
#'}
EVChangeVariableGrid <- function (EVFile, acoVarName, timeDistanceGridType=NULL, depthGridType=NULL, 
                                  timeDistanceGridDistance = NULL, depthGridDistance = NULL, EVLineName = NULL,showenum=FALSE) {
  #enums for grid types
  enumHz=data.frame(enumNbr=0:6,type=c('None','Time','GPS','Vessel log','Pings',
                                       'GPS','Vessel log'),units=c('None','minutes','Nautical mile','Nautical mile',
                                                                   'Pings','m','m'))
  enumVt=data.frame(enumNbr=0:2,type=c('None','depth','reference line'),units=c('none','m','m'))
  if(showenum)
  {
    warning('Only showing Echoview grid enum values. All other function arguments are ignored.')
    return(list(timeDistanceGridenum=enumHz,deothGridenum=enumVt))
  }  
  msgV=paste(Sys.time(),': Starting grid settings inspection/change for',acoVarName)
  
  
  #find acoustic variable:
  acousticVar=EVAcoVarNameFinder(EVFile, acoVarName = acoVarName)
  msgV=c(msgV,acousticVar$msg)
  acousticVar=acousticVar$EVVar
  
  gridObj= acousticVar[["Properties"]][["Grid"]]
  #get old grid values for error checking
  oldDepthLength <-gridObj$DepthRangeSeparation()
  oldDistLength <- gridObj$TimeDistanceSeparation()
  #get old grid types
  oldDepthType <- gridObj$DepthRangeMode()
  oldDistType <- gridObj$TimeDistanceMode()
  
  msg=paste(Sys.time(),' : Time/distance (horizontal settings) prior to change are:',enumHz$type[oldDistType+1],
            'value =',oldDistLength,enumHz$units[oldDistType+1])
  message(msg)
  msgV=c(msgV,msg)
  predataRangeSettings=list(timeDistHzEnum=oldDistType,depthEnum= oldDepthType,timeDistValue=  oldDistLength,
                            depthValue=  oldDepthLength,timeDistHzUnits=as.character(enumHz$units[oldDistType+1]),
                            depthUnits=as.character(enumVt$units[oldDepthType+1]),referenceLineName=NULL,referenceLineObj=NULL)
  refLineText=''
  if(oldDepthType==2) {
    oldEVRefLineObj=gridObj$DepthRangeReferenceLine()
    refLineText=paste('; referenced to line name =',oldEVRefLineObj$Name())
    predataRangeSettings$referenceLineName=oldEVRefLineObj$Name()
    predataRangeSettings$referenceLineObj=oldEVRefLineObj
  }
  msg=paste(Sys.time(),' : Depth (vertical settings) prior to change are:',enumVt$type[oldDepthType+1],
            'value =',oldDepthLength,enumVt$units[oldDepthType+1],refLineText)
  
  message(msg)
  msgV=c(msgV,msg)
  
  
  #catch when the function is used for inspection only
  if(is.null(timeDistanceGridType)& is.null(depthGridType)){
    msg=paste(Sys.time(),' : Both ARGs timeDistanceGridType and depthGridType are NULL so EVChangeVariableGrid() is returning existing grid settings for', acoVarName)
    message(msg)
    msgV=c(msgV,msg)
    return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=NULL,msg=msgV))
  }
  
  #make changes:
  if(timeDistanceGridType< min(enumHz$enumNbr) | timeDistanceGridType>max(enumHz$enumNbr))
  {
    msg=paste(Sys.time(),' : Incorrect time/distance grid type.  ARG timeDistanceGridType > ',
              min(enumHz$enumNbr), 'and < ', max(enumHz$enumNbr)+1)
    warning(msg)
    msgV=c(msgV,msg)
    return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=NULL,msg=msgV))
  }
  if(depthGridType< min(enumVt$enumNbr) | depthGridType>max(enumVt$enumNbr))
  {
    msg=paste(Sys.time(),' : Incorrect depth grid type.  ARG depthGridType > ',
              min(enumVt$enumNbr), 'and < ', max(enumVt$enumNbr)+1)
    warning(msg)
    msgV=c(msgV,msg)
    return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=NULL,msg=msgV))
  }
  #catch type zero:
  if(timeDistanceGridType==0)
  {
    msg=paste(Sys.time(),' : timeDistanceGridType is 0, i.e. no time/horizontal distance grid, setting timeDistanceGridDistance =0')
    message(msg)
    msgV=c(msgV,msg)
    timeDistanceGridDistance=0
  } 
  
  if(depthGridType==0)
  {
    msg=paste(Sys.time(),' : depthGridType is 0, i.e. no depth grid, setting depthGridDistance =0')
    message(msg)
    msgV=c(msgV,msg)
    depthGridDistance=0
  } 
  
  
  #change the horizontal and vertical grids
  horizontal <- gridObj$SetTimeDistanceGrid(timeDistanceGridType, timeDistanceGridDistance)
  vertical   <- gridObj$SetDepthRangeGrid(depthGridType, depthGridDistance)
  
  # if(!horizontal){
  #   msg=paste(Sys.time(),': Failed to set time/distance grid.')
  #   warning(msg)
  #   msgV=c(msgV,msg)
  #   return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=NULL,msg=msgV))
  # }
  # if(!vertical){
  #   msg=paste(Sys.time(),': Failed to set depth grid.')
  #   warning(msg)
  #   msgV=c(msgV,msg)
  #   return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=NULL,msg=msgV))
  # }
  #get new grid values for error checking
  newDepthLength <-gridObj$DepthRangeSeparation()
  newDistLength <- gridObj$TimeDistanceSeparation()
  #get new grid types
  newDepthType <- gridObj$DepthRangeMode()
  newDistType <- gridObj$TimeDistanceMode()
  
  postdataRangeSettings=list(timeDistHzEnum=newDistType,depthEnum= newDepthType,timeDistValue=  newDistLength,
                             depthValue=  newDepthLength,timeDistHzUnits=as.character(enumHz$units[newDistType+1]),
                             depthUnits=as.character(enumVt$units[newDepthType+1]),referenceLineName=NULL,referenceLineObj=NULL)
  
  
  #change the reference line for the depth grid
  if (depthGridType == 2){
    if(is.null(EVLineName)){
      msg=paste(Sys.time(),': ARG depthGrid is set to line referenced grid, but no line has been provided, i.e. ARG EVLineName=NULL')
      warning(msg)
      msgV=c(msgV,msg)
      return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=postdataRangeSettings,msg=msgV))
    }
    newLine=EVFindLineByName(EVFile=EVFile, lineName=EVLineName)
    gridObj[["DepthRangeReferenceLine"]] <- newLine
    if(gridObj[["DepthRangeReferenceLine"]]$Name()==EVLineName) {
      msg=paste(Sys.time(),': Vertical grid line reference line changed to ',EVLineName)
      message(msg)
      msgV=c(msgV,msg)
      postdataRangeSettings$referenceLineName= gridObj[["DepthRangeReferenceLine"]]$Name()
      postdataRangeSettings$referenceLineObj= gridObj[["DepthRangeReferenceLine"]]
    } else
    {
      msg=paste(Sys.time(),': Failed to change vertical grid line reference line changed to ',EVLineName)
      warning(msg)
      msgV=c(msgV,msg)
      return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=postdataRangeSettings,msg=msgV))}
  }
  msg=paste(Sys.time(),': All integration grid setttings successfully updated in acoustic variable',acoVarName)
  message(msg)
  msgV=c(msgV,msg)
  return(list(predataRangeSettings=predataRangeSettings,postdataRangeSettings=postdataRangeSettings,msg=msgV))
}