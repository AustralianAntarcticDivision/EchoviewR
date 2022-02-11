
#' Schools Detection in Echoview
#' 
#' This function performs schools detection in Echoview using COM scripting.
#' @param EVApp =NULL An Echoview Application COM object
#' @param EVFile =NULL An Echoview file COM object
#' @param acoVarName A string containing the name of the acoustic variable to perform the analysis on
#' @param outputRegionClassName A string containing the name of the output region
#' @param deleteExistingRegions Logical TRUE or FALSE 
#' @param distanceMode for schools detection (see Echoview help).
#' @param maximumHorizontalLink The maximum horizontal link in meters 
#' @param maximumVerticalLink The maximum vertical link in meters
#' @param minimumCandidateHeight The minimum candidate height in meters
#' @param minimumCandidateLength the minimum candidate length in meters
#' @param minimumSchoolHeight The minimum school height in meters
#' @param minimumSchoolLength The minimum school length in meters
#' @param dataThreshold minimum integration threshold (units: dB re 1m^-1)
#' @return a list object with four elements. $nbrOfDetectedschools, $thresholdData, $schoolsSettingsData, and $msg: message for processing log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'EVSchoolsDetect(EVFile = EVFile,
#'                acoVarName='120 7x7 convolution',
#'                outputRegionClassName = 'aggregations',
#'                deleteExistingRegions = TRUE,
#'                distanceMode = "GPS distance",
#'                maximumHorizontalLink = 15, #m
#'                maximumVerticalLink = 5,#m
#'                minimumCandidateHeight = 1, #m
#'                minimumCandidateLength = 10, #m
#'                minimumSchoolHeight = 2, #m
#'                minimumSchoolLength = 15, #m
#'                dataThreshold = -80)
#'}


EVSchoolsDetect <- function(
  EVApp=NULL,
  EVFile=NULL,
  acoVarName,
  outputRegionClassName,
  deleteExistingRegions,
  distanceMode,
  maximumHorizontalLink,
  maximumVerticalLink,
  minimumCandidateHeight,
  minimumCandidateLength,
  minimumSchoolHeight,
  minimumSchoolLength,
  dataThreshold) {
  
  #find acoustic variable:
  varObj <- EVAcoVarNameFinder(EVFile = EVFile, acoVarName = acoVarName)
  msgV   <- varObj$msg
  varObj <- varObj$EVVar
  if (is.null(varObj)) {
    msgV <- c(msgV,paste(Sys.time(),' : Stopping schools detection, acoustic variable not found',sep=''))
    message(msgV)
    return(list(nbDetections = NULL, msg = msgV))
  }
  
  #find region class
  regObj <- EVRegionClassFinder(EVFile = EVFile, regionClassName = outputRegionClassName)
  msgV   <- c(msgV,regObj$msg)
  regObj <- regObj$regionClass
  if (is.null(regObj)) {
    msgV <- c(msgV,paste(Sys.time(), ' : Stopping schools detection, region class not found', sep = ''))
    message(msgV[2])
    return(list(nbDetections = NULL, msg = msgV))
  }
  
  #handling exisiting regions:
  if (deleteExistingRegions) {
    msgV <- c(msgV,EVDeleteRegionClass(EVFile = EVFile, regionClassCOMObj = regObj))
  } else {
    msg  <- paste(Sys.time(), ' : adding detected regions those existing in region class ', regObj$Name(), sep = '')
    message(msg)
    msgV <- c(msgV, msg)}
  
  #set threshold
  thresRes <- EVminThresholdSet(varObj = varObj, thres = dataThreshold)
  msgV     <- c(msgV, thresRes$msg)
  
  #set schools detection parameters
  #check EV version
  EV_version=NULL
  if(!is.null(EVApp)) {EV_version=as.numeric(strsplit(EVApp$Version(),'\\.')[[1]][1])
  msg=paste(Sys.time(),": Echoview version is",EV_version)
  message(msg)
  msgV=c(msgV,msg)}
  if(is.null(EV_version)){ #assume if EVApp is not passed in then EV version <12
    schoolDetSet <- EVSchoolsDetSet(EVFile = EVFile, varObj = varObj, distanceMode = distanceMode,
                                    maximumHorizontalLink = maximumHorizontalLink,
                                    maximumVerticalLink = maximumVerticalLink,
                                    minimumCandidateHeight = minimumCandidateHeight,
                                    minimumCandidateLength = minimumCandidateLength,
                                    minimumSchoolHeight = minimumSchoolHeight,
                                    minimumSchoolLength = minimumSchoolLength)}
  if(EV_version>11){
    if(grep('GPS',distanceMode)==1) distanceMode='GPSDistance' #catch the version>11 difference in distance specification
    schoolDetSet <- EVSchoolsDetSetV12(EVApp=EVApp,
      EVFile = EVFile, varObj = varObj, distanceMode = distanceMode,
                                                    maximumHorizontalLink = maximumHorizontalLink,
                                                    maximumVerticalLink = maximumVerticalLink,
                                                    minimumCandidateHeight = minimumCandidateHeight,
                                                    minimumCandidateLength = minimumCandidateLength,
                                                    minimumSchoolHeight = minimumSchoolHeight,
                                                    minimumSchoolLength = minimumSchoolLength)}
    
    msgV <- c(msgV, schoolDetSet$msg)
  msg  <- paste(Sys.time(), ' : Detecting schools in variable ', varObj$name(), sep = '')
  message(msg)
  msgV <- c(msgV,msg)
  nbrDetSchools <- varObj$DetectSchools(regObj$Name())
  if (nbrDetSchools == -1) {
    msg <- paste(Sys.time(),' : Schools detection failed.')
    stop(msg)
  } else {
    msg <- paste(Sys.time(), ' : ',nbrDetSchools , ' schools detected in variable ', 
                 varObj$Name(), sep = '')
    message(msg)}
  msgV <- c(msgV, msg)
  out  <- list(nbrOfDetectedschools = nbrDetSchools, thresholdData = thresRes, 
               schoolsSettingsData = schoolDetSet, msg = msgV)
  return(out)
}
