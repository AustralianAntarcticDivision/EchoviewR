#' Export an Echoview acoustic variable to an image file
#' 
#' This function exports an Echoview acoustic variable to an image file.
#' 
#' @details The following image types are available: *.bmp, *.jpg, *.jpeg, *.png, *.tif, *.tiff. 
#'@param EVFile Echoview file COM object
#'@param acoVarName The acoustic variable name to export
#'@param filePath Path and filename for the resulting image file.  The image type is controlled here by specifying the desired extension, e.g. *.jpg
#'@param heightInPixels Vertical resolution of exported image in pixels.  NB Echoview requires a 25 pixel minimum vertical resolution 
#'@param pingRange =c(-1,1) Optional ping range over which to export the Echoview line object.  Default is to export the lien for all pings
#'@return List of length two [1] Boolean flag indicating success or failure of export, and [1] message vector
#'@export 
EVExportEchogramToImage=function(EVFile,acoVarName,filePath,heightInPixels,pingRange=c(-1,-1))
{
  msgV=paste(Sys.time(),': Attempting to export image to',filePath,'using acoustic variable',acoVarName)
  message(msgV)
  
  if(heightInPixels<25)
  {
    msg=paste(Sys.time(),'Echoview requires a 25 pixel minimum image height.  Changing ARG heightInPixels from',heightInPixels,'to 25')
    warning(msg)
    msgV=c(msgV,msg)
    heightInPixels=25
  }
  AObj=EVAcoVarNameFinder(EVFile=EVFile, acoVarName)
  msgV=c(msgV,AObj$msg)
  AObj=AObj$EVVar
  flag=AObj$ExportEchogramToImage(filePath, heightInPixels, pingRange[1], pingRange[2])
  if(flag)
  {
    flagFileExist=file.exists(filePath)
    if(flagFileExist) { msg=paste(Sys.time(),'Image export successful')
    message(msg)
    msgV=c(msgV,msg)} else {
      msg=paste(Sys.time(),'Echoview is reporting succesful image file export, but the file cannot be found.  Check your directory path.')
     warning(msg)
    msgV=c(msgv,msg)}
    } else {
    msg=paste(Sys.time(), 'Failed to export image file',filePath,'from acoustic variable',acoVarName)
    warning(msg)
    msgV=c(msgV,msg)
    }
    
  invisible(list(flag=flag,msg=msgV))
}


