# Copyright (C) 2015 Lisa-Marie Harrison, Martin Cox, Georg Skaret and Rob Harcourt.
# 
# This file is part of EchoviewR.
# 
# EchoviewR is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# EchoviewR is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with EchoviewR.  If not, see <http://www.gnu.org/licenses/>.
#

#' Open an existing Echoview file (.EV) 
#' 
#' This function opens an existing Echoview (.EV) file using COM scripting.  
#' @param EVAppObj An EV application COM object arising from the call COMCreate('EchoviewCom.EvApplication')
#' @param fileName An Echoview file path and name.
#' @return a list object with two elements.  $EVFile: EVFile COM object, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')
#'}
EVOpenFile <- function (EVAppObj, fileName) {
  nbrFilesOpen <- EVAppObj$EvFiles()$Count()
  msg <- paste(Sys.time(), ' : There are currently ', nbrFilesOpen,  
               ' EV files open in the EV application.', sep = '')
  message(msg[1])
  chkAlreadyOpen <- EVAppObj[['EvFiles']]$FindByFileName(fileName)
  if (!is.null(chkAlreadyOpen)) {
    msg <- c(msg,paste(Sys.time(), ' : File already open. EVOpenFile() 
                       returning existing EV file in EVFile object ', fileName, sep = ''))
    message(msg[2])
    return(list(EVFile = chkAlreadyOpen, msg = msg))
  }
  
  msg <- c(msg,paste(Sys.time(), ' : Opening ', fileName, sep = ''))
  message(msg[2])
  EVFile <- EVAppObj$OpenFile(fileName)
  dF <- EVAppObj$EvFiles()$Count() - nbrFilesOpen #nbr of files now open
  if (dF != 0 | dF != 1) {
    msgTMP <- 'Failed to open EV file, unknown error'
  }   
  if (dF == 1) {
    msgTMP <- 'Opened EV file: '
  }
  if (dF == 0) {
    msgTMP <- 'Check filename. Failed to open EV file: '
  }
  msg <- c(msg, paste(Sys.time(), ' : ', msgTMP, '', fileName, sep = ''))
  message(msg[3])
  invisible(list(EVFile = EVFile, msg = msg))
}


#' Save an open Echoview file (.EV) 
#' 
#' This function saves an existing Echoview (.EV) file using COM scripting.  
#' @param EVFile An Echoview file COM object
#' @return a list object with two elements.  $chk: Boolean check indicating if the file was successfully saved; $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'EVSaveFile(EVFile)
#'}
EVSaveFile <- function (EVFile) {
  chk <- EVFile$Save()
  msg <- paste(Sys.time(), ' : ', ifelse(chk, 'Saved', 'Failed to save'), '', 
               EVFile$FileName(), sep = '')
  if (chk) {
    message(msg)
  } else {
    stop(msg)
  }
  invisible(list(chk = chk, msg = msg))
}


#' Perform save as operation on an open Echoview file (.EV) 
#' 
#' This function performs a save as operation on an existing Echoview (.EV) file using COM scripting.  
#' @param EVFile An Echoview file COM object
#' @param fileName An Echoview file path and name.
#' @return a list object with two elements.  $chk: Boolean check indicating if the file was successfully saved; $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVSaveFile}} \code{\link{EVCloseFile}} 
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'EVSaveAsFile(EVFile = EVFile, fileName = '~~/KAOS/KAOStemplate_test.EV')
#'}
EVSaveAsFile <- function (EVFile,fileName) {
  chk <- EVFile$SaveAs(fileName)
  msg <- paste(Sys.time(), ':', ifelse(chk, 'Saved', 'Failed to save'), '', 
               EVFile$FileName(), sep = '')
  if (chk) {
    message(msg)
  } else {
    stop(msg)
  }
  invisible(list(chk = chk, msg = msg))
}

#' Close an open Echoview file (.EV) 
#' 
#' This function closes an open Echoview file (.EV) via COM scripting
#' @param EVFile An Echoview file COM object
#' @return a list object with two elements.  $chk: Boolean check indicating if the file was successfully closed; $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'EVCloseFile(EVFile)

#'}
EVCloseFile <- function (EVFile) {
  fn  <- EVFile$FileName()
  chk <- EVFile$Close()
  msg <- paste(Sys.time(), ':', ifelse(chk, 'Closed', 'Failed to close'), '', 
               fn, sep = ' ')
  message(msg)
  invisible(list(chk = chk, msg = msg))
}

#' Create a new Echoview file (.EV)
#' 
#' This function creates a new Echoview file (.EV) via COM scripting which may be created from a template file if available
#' @param EVAppObj An EV application COM object arising from the call COMCreate('EchoviewCom.EvApplication')
#' @param templateFn full path and filename for an Echoview template
#' @return a list object with two elements.  $EVFile: EVFile COM object for the newly created Echovie file, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVNewFile(EVAppObj)$EVFile
#'}
EVNewFile <- function (EVAppObj, templateFn = NULL) {
  #ARGS: suggest users specify full path for any template file.
  if (is.null(templateFn)) {
    EvFile <- EVAppObj$NewFile()
    msgV   <- paste(Sys.time(), ':', 'New blank Echoview file created', sep = '')
  }
  if (is.character(templateFn)) {
    msgV <- paste(Sys.time(), ' : Attempting to create an Echoview file from template ', templateFn, sep = '')
    message(msgV)
    EvFile  <- EVAppObj$NewFile(templateFn)
    nbrVars <- EvFile[['Variables']]$Count()
    if (nbrVars == 0) {
      msg <- paste(Sys.time(), ' : Either an incorrect template filename or template contains no variables', sep = '')
      stop(msg)
      msgV <- c(msgV,msg)
    } else {
      msg <- paste(Sys.time(), ' : Echoview file created from template ', 
                   templateFn, sep = '')
      message(msg)
      msgV <- c(msgV, msg)
    }
    
  } 
  if (!is.character(templateFn) & !is.null(templateFn)) {
    msgV <- paste(Sys.time(), ' : Incorrect ARG templateFn specification in EVNewFile()', sep = '')
    stop(msgV)
    return(msgV)}
  invisible(list(EVFile = EvFile, msg = msgV))
}

#' Create a new Echoview fileset
#' 
#' This function creates a new echoview fileset via COM scripting
#' @param EVFile An Echoview file COM object
#' @param filesetName Echoview fileset name to create
#' @return a list object with two elements.  $fileset: created fileset COM object, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVCreateNew}} \code{\link{EVOpenFile}}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'EVCreateFileset(EVFile = EVFile, filesetName = 'example')
#'}
EVCreateFileset <- function (EVFile, filesetName) {
  msgV <- paste(Sys.time(), ' : Creating fileset called ', filesetName, ' in ', 
                EVFile$FileName(), sep = '')
  message(msgV)
  allFilesets <- EVFile[["Filesets"]] 
  chk <- allFilesets$Add(filesetName)
  if (chk) {
    msg <- paste(Sys.time(), ' : Successfully created fileset called ', filesetName, 
                 ' in ', EVFile$FileName(), sep = '')
    message(msg)
    msgV <- c(msgV, msg)
    invisible()
  } 
}

#' Find an Echoview fileset in an Echoview file
#' 
#' This function finds an Echoview fileset in an Echoview file via COM scripting
#' @param EVFile An Echoview file COM object
#' @param filesetName Echoview fileset name to find
#' @return a list object with two elements.  $fileset: found fileset COM object, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVCreateNew}} \code{\link{EVCreateFileset}} \code{\link{EVOpenFile}}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'EVFileset <-EVFindFilesetByName(EVFile = EVFile, filesetName = '038-120-200')$filesetObj
#'}
EVFindFilesetByName <- function (EVFile, filesetName) {
  msgV <- paste(Sys.time(), ' : Searching for fileset name ', filesetName,' in ', 
                EVFile$FileName(), sep = '')
  message(msgV)
  allFilesets <- EVFile[["Filesets"]]
  filesetObj  <- allFilesets$FindByName(filesetName)
  
  if (class(filesetObj)[1] == "COMIDispatch") {
    msg <- paste(Sys.time(), ' : Found the ', filesetObj$Name(), ' fileset in ', 
                 EVFile$FileName(), sep = '')
    message(msg)
    msgV <- c(msgV, msg)
    return(list(filesetObj = filesetObj, msg = msgV))
  } else {
    msg <- paste(Sys.time(), ' : The fileset called ', filesetName, ' not found in ', 
                 EVFile$FileName(), sep = '')
    stop(msg)
    msgV <- c(msgV, msg)
    return(list(filesetObj = NULL, msg = msgV))
  }  
}


#' Add raw data files to an open Echoview file (.EV) 
#' 
#' This function adds raw data files to an open Echoview file (.EV) via COM scripting.  The function assumes the Echoview fileset name already exists.
#' @param EVFile An Echoview file COM object
#' @param filesetName Echoview fileset name
#' @param dataFiles vector of full path and name for each data file.
#' @return a list object with two elements.  $nbrFilesInFileset: number of raw data files in the fileset, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVCreateNew}}
#' @examples
#'\dontrun{
#'filenamesV <- c('~~/KAOS/raw/L0055-D20030115-T171028-EK60.raw', 
#'                  '~~/KAOS/raw/L0055-D20030115-T182914-EK60.raw')
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVNewFile(EVAppObj,templateFn="~~/KAOS/KAOStemplate.EV")$EVFile
#'EVAddRawData(EVFile = EVFile, filesetName = '038-120-200', dataFiles = filenamesV)
#'}
EVAddRawData <- function (EVFile, filesetName, dataFiles) {
  #get number of raw data files currently in "fileset.name" fileset
  destination.fileset <- EVFindFilesetByName(EVFile, filesetName)$fileset
  nbr.of.raw.in.fileset.pre <- destination.fileset[["DataFiles"]]$Count()
  #add new files
  msgV <- paste(Sys.time(), ' : Adding data files to EV file ', sep = '')
  message(msgV)
  for (i in 1:length(dataFiles)) {
    destination.fileset[["DataFiles"]]$Add(dataFiles[i]) 
    msg <- paste(Sys.time(), ' : Adding ', dataFiles[i], ' to fileset name ', 
                 filesetName, sep = '')
    message(msg)
    msgV <- c(msgV, msg)
  }
  
  nbr.of.raw.in.fileset <- destination.fileset[["DataFiles"]]$Count()
  if ((nbr.of.raw.in.fileset - nbr.of.raw.in.fileset.pre) != length(dataFiles)) {
    msg  <- paste(Sys.time(), ' : Number of candidate to number of added file mismatch',
                  sep = '')
    msgV <- c(msgV, msg)
    warning(msg)
  }
  
  invisible(list(nbrFilesInFileset = nbr.of.raw.in.fileset, msg = msgV))
} 


#' Create a new Echoview (.EV) file and adds raw data files to it
#' 
#' This function creates a new Echoview (.EV) file and adds raw data files to it via COM scripting.  Works well when populating an existing Echoview template file with raw data files.  The newly created Echoview file will remain open in Echoview and can be accessed via the $EVFile objected returned by a successful call of this function.
#' @param EVAppObj An EV application COM object arising from the call COMCreate('EchoviewCom.EvApplication')
#' @param EVFileName Full path and filename of Echoview (.EV) file to be created.
#' @param templateFn = NULL Full path and filename of template file if used.
#' @param filesetName Echoview fileset name
#' @param dataFiles vector of full path and name for each data file.
#' @param CloseOnSave = TRUE close the EV file in \code{EVFileName} once saved.
#' @return a list object with two elements.  $EVFile: EVFile COM object for the newly created Echoview file, and $msg: message for processing log. 
#' @details For the example code to run, the example data must be downloaded from.  NB the example code assumes the data and directory structure of the example data has been maintained.
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVAddRawData}}  \code{\link{EVCloseFile}}
#' @examples
#'\dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'pathAndFn=list.files("~~/KAOS/raw/", full.names=TRUE)
#'#remove any evi type files
#'eviLoc=grep('.evi',pathAndFn)
#'if(length(eviLoc)>0) (pathAndFn=pathAndFn[-eviLoc])
#'EVCreateNew(EVAppObj=EVAppObj,
#'                  templateFn="~~/KAOS/KAOStemplate.EV",
#'                  EVFileName='~~/KAOS/kaos.ev',
#'                  filesetName="038-120-200",
#'                  dataFiles=pathAndFn, 
#'                  CloseOnSave = TRUE)
#'}
EVCreateNew <- function (EVAppObj, templateFn = NULL, EVFileName, 
                         filesetName, dataFiles,
                         CloseOnSave=TRUE) {
  msgV <- paste(Sys.time(), ' : Creating new EV file', sep = '')
  message(msgV)
  
  EVFile <- EVNewFile(EVAppObj = EVAppObj, templateFn = templateFn)
  msgV   <- c(msgV, EVFile$msg)
  EVFile <- EVFile$EVFile
  msgV   <- c(msgV, EVAddRawData(EVFile = EVFile, filesetName = filesetName, 
                                 dataFiles = dataFiles)$msg)
  
  msgV   <- c(msgV, EVSaveAsFile(EVFile = EVFile, fileName = EVFileName)$msg)
  if(CloseOnSave)
    msgV=c(msgV,EVCloseFile(EVFile=EVFile)$msg)
  
  return(list(EVFile = EVFile, msg = msgV))
} 


#' Sets minimum data threshold for a variable object
#' 
#' This function sets the minimum data threshold
#' @param varObj An Echoview variable object
#' @param thres The new threshold to be set
#' @return a list object with one element. $thresholdSettings: The new threshold settings
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVAcoVarNameFinder}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, "38 seabed and surface excluded")$EVVar
#' EVminThresholdSet(varObj, -80)
#' }
EVminThresholdSet <- function (varObj, thres) {
  varDat <- varObj[["Properties"]][["Data"]]
  preThresApplyFlag <- varDat$ApplyMinimumThreshold()
  varDat[['ApplyMinimumThreshold']] <- TRUE
  postThresApplyFlag <- varDat$ApplyMinimumThreshold()
  if (postThresApplyFlag) {
    msg <- paste(Sys.time(),' : Apply minimum threshold flag set to TRUE in ',
                 varObj$Name(), sep = '')
    message(msg)
  } else {
    msg <- paste(Sys.time(),' : Failed to set minimum threshold flag in ',
                 varObj$Name(), sep = '')
    stop(msg)
  }
  
  #set threshold value
  preMinThresVal <- varDat$MinimumThreshold()
  varDat[['MinimumThreshold']] <- thres
  postMinThresVal <- varDat$MinimumThreshold()
  if (postMinThresVal == thres) {
    msg2 <- paste(Sys.time(), ' : Minimum threshold successfully set to ', thres, 
                  ' in ', varObj$Name(), sep = '')
    message(msg2)
    msgV <- c(msg, msg2)
  } else {
    msg2 <- paste(Sys.time(), ' : Failed to set minimum threshold in ', 
                  varObj$Name(), sep = '')
    stop(msg2)
    msgV <- c(msg, msg2)
  }
  
  #now try with display threshold
  varDisp <- varObj[["Properties"]][['Display']]
  preDisplayThres <- varDisp$ColorMinimum()
  varDisp[['ColorMinimum']] <- thres
  postDisplayThres <- varDisp$ColorMinimum()
  
  if (thres == postDisplayThres) {
    msg <- paste(Sys.time(), ' : Display threshold also changed to', thres, 
                 ' dB re 1m^-1', sep = '')
    message(msg)
  } else {
    msg <- paste(Sys.time(), ' : Failed to change display threshold to ', thres, 
                 ' dB re 1m^-1 \n', sep = '')
    warning(msg)
  }
  msgV <- c(msg, msgV)
  
  return(list(thresholdSettings = c(preThresApplyFlag = preThresApplyFlag, 
                                    preMinThresVal = preMinThresVal,
                                    postThresApplyFlag = postThresApplyFlag,
                                    preMinThresVal = preMinThresVal,
                                    postMinThresVal = postMinThresVal), msg = msgV))
}  


#' Change schools detection settings
#' 
#' This function changes schools detection settings for an acoustic variable using COM scripting
#' @param EVFile An Echoview file COM object
#' @param varObj the EV acoustic object to change schools detection parameters for
#' @param distanceMode which distance mode to use
#' @param maximumHorizontalLink maximum linking distance for a swarm
#' @param maximumVerticalLink maximum vertical linking distance for a school
#' @param minimumCandidateHeight minimum candidate height
#' @param minimumCandidateLength minimum candidate length
#' @param minimumSchoolHeight minimum school height
#' @param minimumSchoolLength minimum school length
#' @return a list object with two elements. 
#' @keywords Echoview COM scripting
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVAcoVarNameFinder}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, "120 7x7 convolution")$EVVar
#' changeSettings <- EVSchoolsDetSet(EVFile, varObj, distanceMode = "GPS distance",
#'                                 maximumHorizontalLink = 10,
#'                                 maximumVerticalLink = 5,
#'                                 minimumCandidateHeight = 2,
#'                                 minimumCandidateLength = 3,
#'                                 minimumSchoolHeight = 4,
#'                                 minimumSchoolLength = 2)
#'}
EVSchoolsDetSet  <- function (EVFile, varObj, distanceMode,
                              maximumHorizontalLink,
                              maximumVerticalLink,
                              minimumCandidateHeight,
                              minimumCandidateLength,
                              minimumSchoolHeight,
                              minimumSchoolLength) {
  # 20120309 set schools detection parameters.
  #ARGS: EvFile = EV file object;
  # var.nbr = EV virtual variable object number to change the threshold of
  #pars = parameter vector for the schools detection module. vector is structured as follows:
  #     pars[1] = Distance Mode;
  # pars[2] = Maximum horizontal link distance (m);
  #   pars[3] =  Maximum vertical link (m);
  #    pars[4] = MinimumCandidateHeight (m);
  #     pars[5] = MinimumCandidateLength (m);
  #     pars[6] = MinimumSchoolHeight (m);
  #     pars[7] = MinimumSchoolLength (m)
  #returns dataframe of current parameters, revised parameters
  setVec <- c(maximumHorizontalLink, maximumVerticalLink, minimumCandidateHeight, 
              minimumCandidateLength, minimumSchoolHeight,minimumSchoolLength)
  
  if (!all(is.numeric(setVec))) {
    stop('Non-numeric ARG in school detection distance settings')
  }
  
  #get schools object from the current EvFile properties:  
  school.obj <- EVFile[["Properties"]][["SchoolsDetection2D"]] 
  
  #get current school detection parameters
  pre.distmode     <- school.obj[['DistanceMode']]
  pre.maxhzlink    <- school.obj[["MaximumHorizontalLink"]]
  pre.maxvtlink    <- school.obj[["MaximumVerticalLink"]]
  pre.mincandHt    <- school.obj[["MinimumCandidateHeight"]]
  pre.mincandLen   <- school.obj[["MinimumCandidateLength"]]
  pre.minSchoolHt  <- school.obj[["MinimumSchoolHeight"]]
  pre.minSchoolLen <- school.obj[["MinimumSchoolLength"]]
  preSettingDistances <- c(pre.maxhzlink = pre.maxhzlink,
                           pre.maxvtlink = pre.maxvtlink, pre.mincandHt = pre.mincandHt, 
                           pre.mincandLen = pre.mincandLen, pre.minSchoolHt = pre.minSchoolHt, 
                           pre.minSchoolLen = pre.minSchoolLen)
  
  #set school parameters
  school.obj[["DistanceMode"]]           <- distanceMode
  school.obj[["MaximumHorizontalLink"]]  <- maximumHorizontalLink
  school.obj[["MaximumVerticalLink"]]    <- maximumVerticalLink
  school.obj[["MinimumCandidateHeight"]] <- minimumCandidateHeight
  school.obj[["MinimumCandidateLength"]] <- minimumCandidateLength
  school.obj[["MinimumSchoolHeight"]]    <- minimumSchoolHeight
  school.obj[["MinimumSchoolLength"]]    <- minimumSchoolLength
  
  #check settings have been applied by getting current (post change) school detection parameters
  post.distmode     <- school.obj[["DistanceMode"]]
  post.maxhzlink    <- school.obj[["MaximumHorizontalLink"]]
  post.maxvtlink    <- school.obj[["MaximumVerticalLink"]]
  post.mincandHt    <- school.obj[["MinimumCandidateHeight"]]
  post.mincandLen   <- school.obj[["MinimumCandidateLength"]]
  post.minSchoolHt  <- school.obj[["MinimumSchoolHeight"]]
  post.minSchoolLen <- school.obj[["MinimumSchoolLength"]]
  postSettingDistances <- c(post.maxhzlink = post.maxhzlink, post.maxvtlink = post.maxvtlink, 
                            post.mincandHt = post.mincandHt, post.mincandLen = post.mincandLen,
                            post.minSchoolHt = post.minSchoolHt, post.minSchoolLen = post.minSchoolLen)
  
  if (post.distmode != distanceMode) {
    msg <- paste(Sys.time(), " : Failed to set distance mode in schools detection", sep = "")
    invisible(msg)
    stop(msg)
  }
  setDiff <- which(postSettingDistances != setVec)
  if (length(setDiff) > 0) {
    msg <- paste(Sys.time(), ' : Failed to set schools detection parameters: ',
                 paste(names(postSettingDistances)[setDiff], collapse = ', '), sep = '')
    invisible(msg)
    stop(msg)
  } else {
    msg <- paste(Sys.time(), ' : Set schools detection parameters: Distance mode = ', 
                 post.distmode, ' ', paste(names(postSettingDistances),'=', postSettingDistances, collapse = '; '), sep = '')  
    message(msg)
  }
  out <- list(pre.distmode = pre.distmode, preSettingDistances = preSettingDistances,
              post.distmode = post.distmode, postSettingDistances = postSettingDistances, msg = msg)
  return(out)
}


#' Find an acoustic variable by name
#' 
#' This function finds an acoustic variable in an Echoview file by name and 
#' returns the variable pointer.
#' @param EVFile An Echoview file COM object
#' @param acoVarName The name of an acoustic variable in the Echoview file
#' @return a list object with two elements. $EVVar: An Echoview acoustic variable object, and $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, "120 7x7 convolution")$EVVar
#' }
EVAcoVarNameFinder <- function (EVFile, acoVarName) {
  
  obj <- EVFile[["Variables"]]$FindByName(acoVarName)
  if (is.null(obj)) {
    obj <- EVFile[["Variables"]]$FindByShortName(acoVarName)
  }
  if (is.null(obj)) {
    msg <- paste(Sys.time(), ' : Variable not found ', acoVarName, sep = '')
    stop(msg)
    obj <- NULL
  } else {
    msg <- paste(Sys.time(), ' : Variable found ', acoVarName, sep = '') 
    message(msg)
  }
  return(list(EVVar = obj, msg = msg))  
}

#' Find an Echoview region class by name
#' 
#' This function finds an echoview region class by name.
#' @param EVFile An Echoview file COM object
#' @param regionClassName A string containing the name of an Echoview region 
#' @return a list object with two elements. $regionClass: The class of the region, and $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' aggregationsClass <- EVRegionClassFinder(EVFile, "aggregations")$regionClass
#' }
EVRegionClassFinder <- function (EVFile, regionClassName) {
  obj <- EVFile[["RegionClasses"]]$FindByName(regionClassName)
  if (is.null(obj)) {
    msg <- paste(Sys.time(), ' : Region class not found -', regionClassName, sep = '')
    stop(msg)
    obj <- NULL
  } else {
    msg <- paste(Sys.time(), ' : Region class found -', regionClassName, sep = '') 
    message(msg)
  }
  return(list(regionClass = obj, msg = msg))
}

#' Delete an Echoview region class 
#' 
#' This function deletes a region class within an Echoview object using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param regionClassCOMObj An Echoview region object
#' @return a list object with two elements. $EVVar: An Echoview acoustic variable object, and $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVRegionClassFinder}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' exampleClass <- EVRegionClassFinder(EVFile, "003")$regionClass
#' EVDeleteRegionClass(EVFile, exampleClass)
#' }
EVDeleteRegionClass <- function (EVFile, regionClassCOMObj) {
  classChk <- class(regionClassCOMObj)[1]
  if (classChk != 'COMIDispatch') {
    msg <- paste(Sys.time(), ' : attempted to pass non-COM object in ARG regionClassCOMObj', sep = '')
    stop(msg)
  }
  nbrRegionsPre <- EVFile[['Regions']]$Count()
  del <- EVFile[['Regions']]$DeleteByClass(regionClassCOMObj)
  nbrRegionsDel <- nbrRegionsPre - EVFile[['Regions']]$Count()
  if (is.null(del)) {
    msg <- paste(Sys.time(), ' : Regions in region class', regionClassCOMObj$Name(), 
                 ' not deleted', sep = '')
    message(msg)
  } else {
    msg <- paste(Sys.time(), ' : Regions in region class, ', regionClassCOMObj$Name(), 
                 ', deleted. ', nbrRegionsDel , ' individual regions deleted.', sep = '') 
    message(msg)
  }
  invisible(msg)
}

#' Schools Detection in Echoview
#' 
#' This function performs schools detection in Echoview using COM scripting.
#' @param EVFile An Echoview file COM object
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
  EVFile,
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
  
  schoolDetSet <- EVSchoolsDetSet(EVFile = EVFile, varObj = varObj, distanceMode = distanceMode,
                                  maximumHorizontalLink = maximumHorizontalLink,
                                  maximumVerticalLink = maximumVerticalLink,
                                  minimumCandidateHeight = minimumCandidateHeight,
                                  minimumCandidateLength = minimumCandidateLength,
                                  minimumSchoolHeight = minimumSchoolHeight,
                                  minimumSchoolLength = minimumSchoolLength)
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


#' Export integration by regions from an Echoview acoustic variable
#' 
#' This function performs integration by regions and exports the results using COM scripting.
#' @param EVFile An Echoview file object
#' @param acoVarName A string containing the name of an Echoview acoustic variable
#' @param regionClassName A string containing the name of an Echoview region class
#' @param exportFn export filename and path 
#' @param dataThreshold An optional data threshold for export
#' @return a list object with one element, $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVIntegrationByRegionsExport(EVFile, "120 aggregations", "aggregations", exportFn = "~~/KAOS/EVIntegrationByRegionsExport_example.csv")
#' }
EVIntegrationByRegionsExport <- function (EVFile, acoVarName, regionClassName, exportFn,
                                          dataThreshold = NULL) {
  
  acoVarObj <- EVAcoVarNameFinder(EVFile = EVFile, acoVarName = acoVarName)
  msgV      <- acoVarObj$msg
  acoVarObj <- acoVarObj$EVVar
  EVRC      <- EVRegionClassFinder(EVFile = EVFile, regionClassName = regionClassName)
  msgV      <- c(msgV, EVRC$msg)
  RC        <- EVRC$regionClass
  if (is.null(dataThreshold)) {
    msg <- paste(Sys.time(),' : Removing minimum data threshold from ', acoVarName, sep = '')
    message(msg)
    msgV   <- c(msgV,msg)
    varDat <- acoVarObj[["Properties"]][["Data"]]
    varDat[['ApplyMinimumThreshold']] <- FALSE
  } else {
    msg <- EVminThresholdSet(varObj=acoVarObj,thres= dataThreshold)$msg
    message(msg)
    msgV <- c(msgV, msg)
  }
  
  msg <- paste(Sys.time(), ' : Starting integration and export of ', regionClassName, sep = '')
  message(msg)
  success <- acoVarObj$ExportIntegrationByRegions(exportFn, RC)
  
  if (success) {
    msg <- paste(Sys.time(), ' : Successful integration and export of ', regionClassName, sep = '')
    message(msg)
    msgV <- c(msgV, msg)
  } else {
    msg <- paste(Sys.time(), ' : Failed to integrate and/or export ', regionClassName, sep = '')
    warning(msg)
    msgV <- c(msgV, msg)
  } 
  
  invisible(list(msg = msgV))
}


#' Add a calibration file (.ecs) to a fileset
#' 
#' This function adds a calibration file (.ecs) to a fileset using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param filesetName An Echoview fileset name
#' @param calibrationFile An Echoview calibration (.ecs) file path and name
#' @return a list object with one element. $msg: message for processing log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVAddCalibrationFile(EVFile = EVFile, filesetName = '038-120-200', 
#' calibrationFile = '~~/KAOS/20120326_KAOS_SimradEK5.ecs')
#'}

EVAddCalibrationFile <- function (EVFile, filesetName, calibrationFile) {
  
  destination.fileset <- EVFindFilesetByName(EVFile, filesetName)$fileset
  add.calibration <- destination.fileset$SetCalibrationFile(calibrationFile)  
  
  message(paste(Sys.time(), ' : Adding ', basename(calibrationFile),' to fileset name ', filesetName, sep = ''))
  
  if (add.calibration) {
    msg <- paste(Sys.time(), "Success: Added calibration file", basename(calibrationFile), "to fileset name", filesetName)
  } else {
    msg <- paste(Sys.time(), "Error: Could not add calibration file", basename(calibrationFile), "to fileset name", filesetName)
  }
  message(msg)
  invisible(msg)
}

#' Find names of all raw files in a fileset
#' 
#' This function returns the names of all .raw files in a fileset using COM scripting
#' @param EVFile An Echoview file COM object
#' @param filesetName An Echoview fileset name
#' @return A character vector containing the names of all .raw files in the fileset 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' file.names <- EVFilesInFileset(EVFile = EVFile, filesetName = '038-120-200')
#'}

EVFilesInFileset <- function (EVFile, filesetName) {
  
  fileset.loc <- EVFindFilesetByName(EVFile, filesetName)$filesetObj
  nbr.of.raw.in.fileset <- fileset.loc[["DataFiles"]]$Count()
  
  raw.names <- 0
  for (i in 0:(nbr.of.raw.in.fileset - 1)) {
    raw.names[i + 1] <- basename(fileset.loc[["DataFiles"]]$Item(i)$FileName())
  }
  
  message(paste(Sys.time(), ' : Returned names for ', nbr.of.raw.in.fileset, 
                ' data files in fileset ', filesetName ,sep = ''))
  
  return(raw.names)
  
}


#' Clear all files from a fileset
#' 
#' This function clears all .raw files from a fileset using COM scripting
#' @param EVFile An Echoview file COM object
#' @param filesetName An Echoview fileset name
#' @return A list object with one element. $msg message for processing log
#' @keywords Echoview COM scripting 
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVClearRawData(EVFile = EVFile, filesetName = '038-120-200')
#'}

EVClearRawData <- function (EVFile, filesetName) {
  
  destination.fileset   <- EVFindFilesetByName(EVFile,filesetName)$filesetObj
  nbr.of.raw.in.fileset <- destination.fileset[["DataFiles"]]$Count()
  
  #remove files
  msg <- paste(Sys.time(), ' : Removing data files from EV file ', sep ="")
  message(msg)
  
  while (nbr.of.raw.in.fileset > 0) {
    dataFiles <- destination.fileset[["DataFiles"]]$Item(0)$FileName()
    
    rmfile <- destination.fileset[["DataFiles"]]$Item(0)
    destination.fileset[["DataFiles"]]$Remove(rmfile) 
    nbr.of.raw.in.fileset <- destination.fileset[["DataFiles"]]$Count()
    
    msg <- paste(Sys.time(), ' : Removing ', basename(dataFiles),' from fileset name ', 
                 filesetName, sep = "")
    message(msg)
  }
  
  invisible(msg)
  
}

#' Find the time and date of the start and end of an Echoview fileset
#'
#' This function finds the date and time of the first and last measurement in a fileset using COM scripting
#' @param EVFile An Echoview file COM object
#' @param filesetName An Echoview fileset name
#' @return A list object with two elements $start.time: The date and time of the first measurement in the fileset, and $end.time: The date and time of the last measurement in the fileset
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' survey.time = EVFindFilesetTime(EVFile = EVFile, filesetName = '038-120-200')
#' start.time <- survey.time$start.time
#' end.time <- survey.time$end.time
#'}

EVFindFilesetTime <- function (EVFile, filesetName) {
  
  fileset.loc <- EVFindFilesetByName(EVFile, filesetName)$filesetObj
  
  #find date and time for first measurement
  start.date          <- as.Date(trunc(fileset.loc$StartTime()), origin = "1899-12-30")
  percent.day.elapsed <- fileset.loc$StartTime() - trunc(fileset.loc$StartTime())
  seconds.elapsed     <- 86400 * percent.day.elapsed
  start.time          <- as.POSIXct(seconds.elapsed, origin = start.date, tz = "GMT")
  
  #find date and time for last measurement
  end.date            <- as.Date(trunc(fileset.loc$EndTime()), origin = "1899-12-30")
  percent.day.elapsed <- fileset.loc$EndTime() - trunc(fileset.loc$EndTime())
  seconds.elapsed     <- 86400 * percent.day.elapsed
  end.time            <- as.POSIXct(seconds.elapsed, origin = end.date, tz = "GMT")
  
  return(list(start.time = start.time, end.time = end.time))
  
}


#' Create a new Echoview region class
#' 
#' This function creates a new region class using COM scripting
#' @param EVFile An Echoview file COM object
#' @param className The name of the new Echoview region class
#' @return A list object with one element. $msg message for processing log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVNewRegionClass(EVFile = EVFile, className = 'test_class')
#'}

EVNewRegionClass <- function (EVFile, className) {
  
  
  for (i in 1:length(className)) {
    
    add.class <- EVFile[["RegionClasses"]]$Add(className[i])
    
    if (add.class) {
      msg <- paste(Sys.time(), ' : Added region class', className[i], 'to EVFile' , sep = ' ')
    }  else {
      msg <- paste(Sys.time(), ' : Error: could not add region class', className[i],'to EVFile' , sep = ' ')
    }
    message(msg)
  }
}

#' Export Sv data for an Echoview acoustic variable by region
#' 
#' This function exports the Sv values as a .csv file for an acoustic variable by region using COM scripting
#' @param EVFile An Echoview file COM object
#' @param variableName Echoview variable name for which to extract the data
#' @param regionName Echoview region name for which to extract the data
#' @param filePath File path and name (.csv) to save the data 
#' @return A list object with one element. $msg message for processing log
#' @keywords Echoview COM scripting 
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVImportRegionDef}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVImportRegionDef(EVFile = EVFile, evrFile = '~~/KAOS/off transect regions/20030114_1200000000.evr', regionName = 'region_1')
#' EVExportRegionSv(EVFile = EVFile, variableName = '120 seabed and surface excluded', regionName = 'region_1', filePath = '~~/KAOS/EVExportRegionSv_example.csv')
#'}

EVExportRegionSv <- function (EVFile, variableName, regionName, filePath) {
  
  acoustic.var <- EVFile[["Variables"]]$FindByName(variableName)
  ev.region    <- EVFile[["Regions"]]$FindByName(regionName)
  export.data  <- acoustic.var$ExportDataForRegion(filePath, ev.region)
  
  if (export.data) {
    msg <- paste(Sys.time(), ' : Exported data for Region ', regionName, ' in Variable ', 
                 variableName, sep = '')
  } else {
    msg <- paste(Sys.time(), ' : Failed to export data' , sep = "")
  }
  
  message(msg)
  invisible(msg)
  
}

#' Change the data range bitmap of an acoustic object
#'
#' This function changes the data range in an Echoview data range bitmap virtual variable
#' @param varObj An Echoview acoustic variable COM object, perhaps resulting from a call of EVAcoVarNameFinder()
#' @param minRng the minimum data value to set
#' @param maxRng the maximum data value to set
#' @return a vector of pre- and post-function call data range settings
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVAcoVarNameFinder}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, acoVarName = "38 data range bitmap")$EVVar
#' EVAdjustDataRngBitmap(varObj, minRng = -90, maxRng = 0)
#'}

EVAdjustDataRngBitmap <- function (varObj, minRng, maxRng) {
  
  message(paste(Sys.time(), " : Adjusting the data range of ", varObj$Name(), sep = ' '))
  
  #check if varObj is an acoustic variable
  if ((class(varObj$AsVariableAcoustic()) == "COMIDispatch") == FALSE) {
    stop(paste(Sys.time(), ' : STOPPED. Input acoustic variable object ', varObj$Name(), ' is not an acoustic variable', sep = ''))
  } 
  
  #get pre-change min and max ranges
  rngAttrib   <- varObj[["Properties"]][["DataRangeBitmap"]]
  preMinrange <- rngAttrib$RangeMinimum()
  preMaxrange <- rngAttrib$RangeMaximum()
  message(paste(Sys.time(),' : Preset data range values; minimum = ', preMinrange, ' maximum =', preMaxrange, sep = ' '))
  
  #change data range
  postMinrangeFlag = rngAttrib[['RangeMinimum']] <- minRng
  postMaxrangeFlag = rngAttrib[['RangeMaximum']] <- maxRng
  
  #check post min and max values
  postMinrange <- rngAttrib$RangeMinimum()
  postMaxrange <- rngAttrib$RangeMaximum()
  
  #create data range values output object:
  datarange <- data.frame(preMinrange = preMinrange, preMaxrange = preMaxrange, 
                          postMinrange = postMinrange, postMaxrange = postMaxrange)
  row.names(datarange) <- varObj$Name()
  
  #check post range set values equal ARGS minRng,maxRng:
  if (postMinrange != minRng | postMaxrange != maxRng) {
    message(paste(Sys.time()," : FAILED to set data range bitmap values in ",varObj$Name(), ' Current data range values are: min =', postMinrange,'; ', postMaxrange))
    
  } else {
    message(paste(Sys.time(), ' : SUCCESS. Data range values in ', varObj$Name(), '; minimum = ', preMinrange,' maximum=', preMaxrange, sep = ' '))
  }
  
  return(datarange = datarange)
} 

#' Find an EV Line object by name
#'
#' This function finds an EV Line in an EV file object by name using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param lineName a string containing the name of the line to find
#' @return an Echoview line object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVLine <- EVFindLineByName(EVFile = EVFile, lineName = "Fixed depth 250 m")
#'}
EVFindLineByName <- function (EVFile, lineName) {
  
  EVLine <- EVFile[["Lines"]]$FindByName(lineName)
  
  #check whether line was found
  if (is.null(EVLine)) {
    message(paste(Sys.time(), "Error: Cannot find a line named", lineName))
  } else {
    message(paste(Sys.time(), "Success: Found line named", lineName))
  }
  
  return(EVLine)
}


#' Change the grid of an acoustic variable
#'
#' This function sets the grid separation and depth reference line for an acoustic variable using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param acousticVar an EV acoustic variable object
#' @param verticalType 0 = no grid, 1 = time (minutes), 2 = GPS distance (NMi), 3 = Vessel Log Distance (NMi), 4 = Pings, 5 = GPS distance (m), 6 = Vessel Log Distance (m). 
#' @param horizontalType 0 = no grid, 1 = depth grid, 2 = use reference line.
#' @param verticalDistance vertical grid line spacing. Not needed if verticalType = 0. 
#' @param horizontalDistance horizontal grid line spacing. Not needed if horizontalType = 0. 
#' @param EVLine an EV line object. Not needed if horizontalType = 0.
#' @return a list object with two elements.  $dataRangeSettings: a vector of pre- and post-function call data range settings, and $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVAcoVarNameFinder}} \code{\link{EVFindLineByName}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' varObj <- EVAcoVarNameFinder(EVFile, acoVarName = "38 seabed and surface excluded")$EVVar
#' EVLine <- EVFindLineByName(EVFile = EVFile, lineName = "Fixed depth 250 m")
#' 
#' #Change grid to 100m vertical distance and 10m depth grid relative to 100m line 
#' EVChangeVariableGrid(EVFile = EVFile, acousticVar = varObj, 
#' verticalType = 5, horizontalType = 2, verticalDistance = 100, horizontalDistance = 10, EVLine)
#' #remove horizontal and vertical grid
#' EVChangeVariableGrid(EVFile = EVFile, acousticVar = varObj, 
#' verticalType = 0, horizontalType = 0)

#'}


EVChangeVariableGrid <- function (EVFile, acousticVar, verticalType, horizontalType, verticalDistance = 0, horizontalDistance = 0, EVLine = NULL) {
  
  #get old grid values for error checking
  old_depth <- acousticVar[["Properties"]][["Grid"]]$DepthRangeSeparation()
  old_distance <- acousticVar[["Properties"]][["Grid"]]$TimeDistanceSeparation()
  
  #change the horizontal and vertical grids
  horizontal <- acousticVar[["Properties"]][["Grid"]]$SetDepthRangeGrid(horizontalType, horizontalDistance)
  vertical   <- acousticVar[["Properties"]][["Grid"]]$SetTimeDistanceGrid(verticalType, verticalDistance)
  
  #change the reference line for the depth grid
  try(if (horizontalType == 2) {
    acousticVar[["Properties"]][["Grid"]][["DepthRangeReferenceLine"]] <- EVLine
  }, silent = TRUE)
  
  
  #get unit type depending on specified grid types
  vertical.codes <- c(0:6)
  vertical.unit.types <- c("no grid", "minutes", "nautical miles", "nautical miles", "pings", "meters", "meters")
  vertical.units <- vertical.unit.types[which(vertical.codes == verticalType)]
  
  if (horizontalType == 0) {
    horizontal.units <- "no grid"
  } else {
    horizontal.units <- "meters"
  }
  
  
  if (vertical) {
    msgv <- paste(Sys.time(), " Changed depth grid separation to ", verticalDistance, vertical.units, sep = "")
  } else {
    if (old_distance == verticalDistance) {
      msgv <- paste(Sys.time(), " Error: Depth grid is already set to ", verticalDistance, vertical.units, sep = "")
      
    } else {
      msgv <- paste(Sys.time(), " Error: Failed to change depth grid separation")
    }
  }
  
  if (horizontal) {
    msgh <- paste(Sys.time(), " Changed horizontal grid line distance to ", horizontalDistance, horizontal.units, sep = "")
  } else {
    if (old_depth == horizontalDistance) {
      msgh <- paste(Sys.time(), " Error: Horizontal grid is already set to ", horizontalDistance, horizontal.units, sep = "")
      
    } else {
      msgh <- paste(Sys.time(), " Error: Failed to change horizontal grid separation")
    }
  }
  
  message(msgv)
  message(msgh)
}

#' Export integration by cells for an acoustic variable
#'
#' This function exports the integration by cells for an acoustic variable using COM scripting. Note: This function will only work if the acoustic variable has a grid.
#' @param EVFile An Echoview file COM object
#' @param variableName a string containing the name of an EV acoustic variable
#' @param filePath a string containing the file path and name to save the exported data to 
#' @return a list object with 1 element: message for progessing log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVExportIntegrationByCells(EVFile = EVFile, 
#' variableName = '38 seabed and surface excluded', 
#' filePath = '~~/KAOS/EVExportIntegrationByCells_example.csv')
#'}


EVExportIntegrationByCells <- function (EVFile, variableName, filePath) {
  
  acoustic.var <- EVFile[["Variables"]]$FindByName(variableName)
  
  #check that the acoustic variable exists
  if (is.null(acoustic.var)) {
    msg <- paste(Sys.time(), "Error:", variableName, "is not an acoustic variable", sep = " ")
    
  } else {
    
    export.data <- acoustic.var$ExportIntegrationByCellsAll(filePath)
    
    if (export.data) {
      msg <- paste(Sys.time(), 'Success: Exported integration by cells for variable', variableName, sep = " ")
    } else {
      msg <- paste(Sys.time(), 'Error: Failed to export data')
    }
  }
  
  message(msg)
  invisible(msg)
  
}

#' Add a new acoustic variable
#'
#' This function adds a new acoustic variable using COM scripting
#' @param EVFile An Echoview file COM object
#' @param oldVarName a string containing the name of the acoustic variable to base the new variable on
#' @param enum Enum code for operator. See Echoview help file on EOperator for enum codes.
#' @return an object: returns the new variable
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' #create a 7x7 convolution of a variable
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVNewAcousticVar(EVFile = EVFile, oldVarName = "38 seabed and surface excluded", enum = 43)
#'}
EVNewAcousticVar <- function (EVFile, oldVarName, enum) {
  
  acoustic.var <- EVFile[["Variables"]]$FindByName(oldVarName)
  
  if (is.null(acoustic.var)) {
    stop(paste(Sys.time(), "Error: Could not find the variable ", oldVarName))    
  } else {
    message(paste(Sys.time(), "Found the variable", oldVarName, sep = " "))
  } 
  
  newVar <- acoustic.var$AddVariable(enum)
  
  if (is.null(newVar)) {
    message(paste(Sys.time(), "Error: Failed to create the new variable"))
  } else {
    message(paste(Sys.time(), "Success: New variable created"))
  } 
  
  return(newVar)
}

#' Change the depth of an Echoview Region
#'
#' This function shifts the depth of an echoview region using COM scripting. Vertical size of the region and depth offset can be changed.
#' @param EVFile An Echoview file COM object
#' @param regionName a string containing the name of the Echoview region
#' @param depthMultiply a numeric value to multiply the vertical size of the region by 
#' @param depthAdd a numeric value to offset the region depth by. Positive = decrease depth; Negative = increase depth
#' @return a list object with one element, $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' #Double region vertical size and decrease depth by 100m
#' EVShiftRegionDepth(EVFile, "testregion", 2, 100)
#' 
#' #Triple region vertical size without changing depth offset
#' EVShiftRegionDepth(EVFile, "testregion", 3, 0)
#' 
#' #Change region depth offset by -50m without changing vertical size
#' EVShiftRegionDepth(EVFile, "testregion", 1, -50)
#'}
EVShiftRegionDepth <- function (EVFile, regionName, depthMultiply, depthAdd) {
  
  region <- EVFile[["Regions"]]$FindByName(regionName)
  
  if (is.null(region)) {
    stop(paste(Sys.time(), "Error: Could not find the region", regionName, sep = " "))
  } else {
    message(paste(Sys.time(), "Found the region", regionName, sep = " "))
  }
  
  shift.depth <- region$ShiftDepth(depthMultiply, depthAdd)
  
  if (shift.depth) {
    msg <- paste("Success: Multiplied", regionName, "depth by", depthMultiply, "and added", depthAdd, "m", sep = " ")
  } else {
    msg <- paste("Error: Failed to change", regionName, "depth", sep = " ")
  }
  
  message(msg)
  return(list(msg=msg))
}


#' Change the time of an Echoview Region
#'
#' This function shifts the time of an echoview region using COM scripting
#' @param EVFile An Echoview file COM object
#' @param regionName a string containing the name of the Echoview region
#' @param days an integer value specifying days to add (positive) or subtract (negative). Default = 0
#' @param hours an integer value specifying hours to add (positive) or subtract (negative). Default = 0
#' @param minutes an integer value specifying minutes to add (positive) or subtract (negative). Default = 0
#' @param seconds an integer value specifying seconds to add (positive) or subtract (negative). Default = 0
#' @param milliseconds an integer value specifying milliseconds to add (positive) or subtract (negative). Default = 0
#' @return a list object with one element, $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' #Shift the region time by 10 seconds
#' EVShiftRegionTime(EVFile, "testregion", seconds = 10)
#' 
#' #Subtract 1 hour from the region time
#' EVShiftRegionTime(EVFile, "testregion", hours = -1)
#'}
EVShiftRegionTime <- function (EVFile, regionName, days = 0, hours = 0, minutes = 0, seconds = 0, milliseconds = 0) {
  
  
  region <- EVFile[["Regions"]]$FindByName(regionName)
  
  if (is.null(region)) {
    stop(paste(Sys.time(), "Error: Could not find the region", regionName, sep = " "))
  } else {
    message(paste(Sys.time(), "Found the region", regionName, sep = " "))
  }
  
  shift.depth <- region$ShiftTime(days, hours, minutes, seconds, milliseconds)
  
  if (shift.depth) {
    msg <- paste("Success: Changed", regionName, "time by", days, "days", hours, "hours", minutes, "mins", seconds, "seconds", "and", milliseconds, "milliseconds", sep = " ")
  } else {
    msg <- paste("Error: Failed to change", regionName, "time", sep = " ")
  }
  
  message(msg)
  invisible(list(msg=msg))
}


#' Gets the calibration file name of a fileset
#' 
#' This function gets the calibration file name of a filesset using COM scripting
#' @param EVFile An Echoview file COM object
#' @param filesetName a string containing the name of the Echoview fileset
#' @return an object: returns the calibration file name
#' @keywords Echoview COM scripting
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVGetCalibrationFileName(EVFile = EVFile, filesetName = "038-120-200")
#'}
EVGetCalibrationFileName <- function (EVFile, filesetName) {
  
  
  fileset <- EVFile[["Filesets"]]$FindByName(filesetName)
  
  if (is.null(fileset)) {
    stop(paste(Sys.time(), "Error: Couldn't file the fileset", filesetName, sep = " "))
  } else {
    message(paste(Sys.time(), "Found the fileset", filesetName, sep = " "))
  }
  
  calibration.file <- fileset$GetCalibrationFileName()
  
  return(calibration.file)
  
}

#' Creates a new line relative region in the current variable
#'
#' This function creates a new line relative region in the current variable using COM scripting. The upper and lower depths are specified using Echoview line objects (these must already exist). Left and right bounds are optionally specified using ping number.
#' @param EVFile An Echoview file COM object
#' @param varName a string containing the name of the acoustic variable to create the region in
#' @param regionName a string containing the name to assign to the new region
#' @param line1 a string containing the name of the upper line limit
#' @param line2 a string containing the name of the lower line limit
#' @param firstPing an optional integer for the ping to begin the region at
#' @param lastPing an optional integer for the ping to end the region at
#' @return returns the EV Region object
#' @keywords Echoview COM scripting
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @export
#' @examples
#' \dontrun{
#'EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#'EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#'
#'#create a region between pings 1 - 100 and depths 20-250m
#'newRegion <- EVNewLineRelativeRegion(EVFile, 
#'"38 seabed and surface excluded", "test", 
#'"Fixed depth 6 m", "Fixed depth 250 m", 1, 100)
#'
#'#create an unbounded region between depths 250-750m
#'newRegion <- EVNewLineRelativeRegion(EVFile, "38 seabed and surface excluded", "test", "Fixed depth 6 m", "Fixed depth 250 m")
#'}
EVNewLineRelativeRegion <- function (EVFile, varName, regionName, line1, line2, firstPing = NA, lastPing = NA) {
  
  acoustic.var <- EVFile[["Variables"]]$FindByName(varName)
  
  if (is.null(acoustic.var)) {
    stop(paste(Sys.time(), "Error: Could not find the variable ", varName))    
  } else {
    message(paste(Sys.time(), "Found the variable", varName, sep = " "))
  } 
  
  line1.obj <- EVFindLineByName(EVFile, line1)
  line2.obj <- EVFindLineByName(EVFile, line2)
  
  new.region <- acoustic.var$CreateLineRelativeRegion(regionName, line1.obj, line2.obj, firstPing, lastPing)
  
  if (is.null(new.region)) {
    stop(paste(Sys.time(), "Error: Unable to create region"))
  }
  else {
    message(paste(Sys.time(), "Success: Line relative region created"))
  }
  
  return(new.region)
  
}


#' Creates a new fixed depth Echoview line
#'
#' This function creates a new Echoview line at a fixed depth using COM scripting
#' @param EVFile An Echoview file COM object
#' @param depth an integer specifying the fixed depth of the new line
#' @param lineName a string containing the name for the new line
#' @return returns the EV Line object
#' @keywords Echoview COM scripting
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @export
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' #create a new line at 50m depth named "testline"
#' newLine <- EVNewFixedDepthLine(EVFile = EVFile, depth = 50, lineName = "testline")
#'}
EVNewFixedDepthLine <- function (EVFile, depth, lineName) {
  
  #check if there is already a line with that name
  line.check <- EVFindLineByName(EVFile, lineName)
  
  if (is.null(line.check) == FALSE) {
    stop(paste(Sys.time(), "Error: A line already exists with the name", lineName, ". New line not created", sep = " "))
  }
  
  new.line <- EVFile[["Lines"]]$CreateFixedDepth(depth)
  new.line[["Name"]] <- lineName
  
  if (new.line$Name() == lineName) {
    message(paste(Sys.time(), "Success: New line", lineName, "created at depth", depth, "m", sep = " "))
  } else {
    stop(paste(Sys.time(), "Error: New line", lineName, "not created", sep = " "))
  }
  
  return(new.line)
  
}


#' Deletes an Echoview line object
#'
#' This function deletes an Echoview line object using COM scripting
#' @param EVFile An Echoview file COM object
#' @param evLine an Echoview line object
#' @return a list object with one element- function message for log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVNewFixedDepthLine}} \code{\link{EVFindLineByName}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' testline <- EVNewFixedDepthLine(EVFile = EVFile, depth = 50, lineName = "test_line")
#' EVDeleteLine(EVFile = EVFile, evLine = testline)
#'}
EVDeleteLine <- function (EVFile, evLine) {
  
  delete.line <- EVFile[["Lines"]]$Delete(evLine)
  
  if (delete.line){
    msg=paste(Sys.time(), "Success: Line deleted")
    message(msg)
  }
  else {
    msg=paste(Sys.time(), "Error: Line not deleted")
    message(msg)}
  invisible(list(msg=msg))
  
}


#' Renames an Echoview Line object
#'
#' This function renames an Echoview line object using COM scripting
#' @param EVFile An Echoview file COM object
#' @param evLine an Echoview line object
#' @param newName a string containing the new name for the line
#' @return a list object with one element- fucntion message for log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVNewFixedDepthLine}} \code{\link{EVFindLineByName}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' testline <- EVNewFixedDepthLine(EVFile = EVFile, depth = 50, lineName = "testline")
#' EVRenameLine(EVFile = EVFile, evLine = testline, newName = "line40")
#'}
EVRenameLine <- function (EVFile, evLine, newName) {
  
  #check if there is already a line with the new name
  line.check <- EVFile[["Lines"]]$FindByName(newName)
  
  if (is.null(line.check) == FALSE) {
    renameFlag <- FALSE
    message(paste(Sys.time(), "Error: A line already exists with the name", newName, ". Line not renamed", sep = " "))
  } else {
    
    evLine[["Name"]] <- newName
    renameFlag = (evLine$Name() == newName)
    
    
    #check that line has been renamed
    if (renameFlag) {
      msg <- (paste(Sys.time(), "Success: Line renamed as", newName, sep = " "))
      message(msg)
    } else {
      message(paste(Sys.time(), "Error: Could not rename line"))
    }
  }
  return(renameFlag)
}

#' Finds an Echoview region by name
#'
#' This function finds an Echoview region by name using COM scripting
#' @param EVFile An Echoview file COM object
#' @param regionName a string containing the name of the region to find
#' @return the Echoview region object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' testRegion <- EVFindRegionByName(EVFile, "Region1")
#'}
EVFindRegionByName <- function (EVFile, regionName) {
  
  ev.region <- EVFile[["Regions"]]$FindByName(regionName)
  
  if (is.null(ev.region)) {
    stop(paste(Sys.time(), "Could not find a region named", regionName, sep = " "))
  } else {
    message(paste(Sys.time(), "Success: Found the region named", regionName, sep = " "))
  }
  
  return(ev.region)
  
}


#' Exports an Echoview region definition
#'
#' This function exports a single region's definition as a .csv file using COM scripting
#' @param EVFile An Echoview file COM object
#' @param regionName a string containing the region name to export definitions for
#' @param filePath a string containing the name and file path of the file to export to
#' @return a list object with one element- fucntion message for log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} 
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' EVExportRegionDef(EVFile, regionName = "Region1", filePath = "~~/KAOS/EVExportRegionDef.csv")
#'}
EVExportRegionDef <- function (EVFile, regionName, filePath) {
  
  ev.region <- EVFindRegionByName(EVFile, regionName)
  
  if (exists("ev.region") == FALSE) {
    stop(paste("Couldn't find region named", regionName, sep = " "))
  }
  
  export.def <- ev.region$ExportDefinition(filePath)
  
  if (export.def) {
    msg=paste(Sys.time(), "Success: Exported region definitions")
    message(msg)
  } else {
    msg=paste(Sys.time(), "Error: Failed to export region definitions")
    warning(msg)
  }
  invisible(list(msg=msg))
}


#' Exports definitions for all Echoview regions in a region class
#'
#' This function exports definitions for all Echoview regions within a region class using COM scripting
#' @param evRegionClass an Echoview region class object
#' @param filePath a string containing the name and file path of the file to export to
#' @return a list object with one element- fucntion message for log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVRegionClassFinder}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' regionClass <- EVRegionClassFinder(EVFile, "aggregations")$regionClass
#' EVExportRegionDefByClass(evRegionClass = regionClass, filePath = "~~/KAOS/EVExportRegionDefByClass_example.csv")
#'}
EVExportRegionDefByClass <- function (evRegionClass, filePath) {
  
  export.def <- evRegionClass$ExportDefinitions(filePath)
  
  if (export.def) {
    msg=paste(Sys.time(), "Success: Exported region definitions")
    message(msg)
  } else {
    message(paste(Sys.time(), "Error: Failed to export region definitions"))
  }
  invisible(list(msg=msg))
}

#' Find the class of an Echoview region object
#'
#' This function finds the class of an Echoview region object using COM scripting.
#' @param evRegion and Echoview Region object
#' @return a string containing the class of the region
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}} \code{\link{EVFindRegionByName}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' 
#' ev.region <- EVFindRegionByName(EVFile, "Region1")
#' EVFindRegionClass(ev.region)
#' }


EVFindRegionClass <- function (evRegion) {
  
  region.class <- evRegion$RegionClass()$Name()
  
  return(region.class)
  
}


#' Generate a coordinate list for a regular rectangular survey 
#'
#'The coordinate list is generated in degrees decimal degree format (dd.ddd), with 
#'Southern hemisphere denoted by negative numbers. Transect length and inter-transect 
#'spacing are specified in km and bearings in degress where North 0 deg, East 90 deg, 
#'South 180 deg and West 270 deg.
#'
#' @param startLon start longitude of survey.
#' @param startLat start latitude of survey.
#' @param lineLengthkm transect line length in km.
#' @param lineSpacingkm inter-transect spacing in km.
#' @param startBearingdeg Orientation of survey grid in degrees.
#' @param numOfLines Number of transects.
#' @return Geographical coordinate list of start and end of line positions
#' @export
#' @seealso \code{\link{zigzagSurvey}}
#' @author Martin Cox \email{martin.cox@@aad.gov.au}
#' @examples
#' \dontrun{
#' (coords=lawnSurvey(startLon=-170,startLat=-60,lineLengthkm=2,lineSpacingkm=0.5,
#'startBearingdeg=30,numOfLines=5))
#'plot(0,0,xlim=range(coords[,1]),
#'ylim=range(coords[,2]),type='n',xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'legend('topright',c('Beginning','End'),col='blue',pch=c(17,15))
#'}
lawnSurvey=function(startLon,startLat,lineLengthkm,lineSpacingkm,startBearingdeg,numOfLines)
{
  lineLengthm=lineLengthkm*1e3;lineSpacingm=lineSpacingkm*1e3
  direct=startBearingdeg+90
  direct[direct>360]=direct-360
  dir2=startBearingdeg+180
  dir2[dir2>360]=dir2-360
  lineBearings=c(startBearingdeg,dir2)
  out=matrix(NA,nrow=2*numOfLines,ncol=2,
             dimnames=list(paste(rep(c('SOL','EOL'),numOfLines),
                                 'line',sort(rep(1:numOfLines,2)),sep=''),
                           c('lon','lat')))
  out[1,]=c(startLon,startLat)
  out[2,]= destPoint(p=out[1,],b=startBearingdeg,d=lineLengthm) 
  if(numOfLines>1){
    for(i in 1:(numOfLines-1))
    {  
      out[(i*2)+1,]=destPoint(p=out[i*2,],b=direct,d=lineSpacingm)
      out[(i*2)+2,]=destPoint(p=out[(i*2)+1,],b=lineBearings[ifelse(i%%2==1,2,1)],d=lineLengthm) 
    }
  }
  return(out)
}


#' Generate a coordinate list for a zig-zag survey 
#'
#'The coordinate list is generated in degrees decimal degree format (dd.ddd), with 
#'Southern hemisphere denoted by negative numbers. Transect length are specified in km 
#'and bearings in degress where North 0 deg, East 90 deg, South 180 deg and West 270 deg.
#'
#' @param startLon start longitude of survey.
#' @param startLat start latitude of survey.
#' @param lineLengthkm transect line length in km.
#' @param startBearingdeg Bearing of each transect in degrees.
#' @param rotationdeg rotation angle for entire survey pattern.
#' @param numOfLines Number of transects.
#' @param proj4string projection string of class \link{CRS-class}
#' @param unrotated \code{FALSE} return rotated coordinates \code{TRUE} list of rotated and unrotated coordinates.
#' @return Geographical coordinates of start and end of line positions.  \code{unrotated=TRUE} list of rotated and unrotated coordinates
#' @export
#' @seealso \link{zigzagSurvey}
#' @author Martin Cox \email{martin.cox@@aad.gov.au}
#' @examples
#' \dontrun{
#' coords=zigzagSurvey(startLon=-100,startLat=-60,lineLengthkm=2,
#' startBearingdeg=30,
#' rotationdeg=10,numOfLines=11)
#'plot(0,0,xlim=range(coords[,1]),ylim=range(coords[,2]),type='n',
#'xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'legend('topright',c('Beginning','End'),col='blue',pch=c(17,15))
#'#Use unrotated=TRUE and check coordinates in the- 
#'#-returned coordinate list object are identical:
#'coordL= zigzagSurvey(startLon=-100,startLat=-60,lineLengthkm=2,
#'startBearingdeg=30,
#' rotationdeg=0,numOfLines=11,unrotated=TRUE)
#' identical(coordL[[1]],coordL[[2]])
#' #display rotated and unrotated coordinates"
#'coordL= zigzagSurvey(startLon=-100,startLat=-60,lineLengthkm=2,
#'startBearingdeg=30,
#' rotationdeg=5,numOfLines=11,unrotated=TRUE) 
#' coords=coordL$unrotatedGeogs; coordsRotate=coordL$rotatedGeogs
#' plot(0,0,xlim=range(c(coords[,1],coordsRotate[,1])),
#' ylim=range(c(coords[,2],coordsRotate[,2])),type='n',
#' xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'arrows(x0=coordsRotate[1:(nrow(coordsRotate)-1),1], 
#'y0=coordsRotate[1:(nrow(coordsRotate)-1),2], 
#'       x1 = coordsRotate[2:nrow(coordsRotate),1], 
#'       y1 = coordsRotate[2:nrow(coordsRotate),2],col='blue')
#'legend('bottomleft',c('Unrotated','Rotated'),lty=1,col=c(1,'blue'))
#'}
zigzagSurvey=function(startLon,startLat,lineLengthkm,startBearingdeg,rotationdeg,numOfLines,
                      proj4string=CRS("+proj=longlat +datum=WGS84"),unrotated=FALSE)
{
  lineLengthm=lineLengthkm*1e3
  angs=c(startBearingdeg,180-startBearingdeg)
  lineIND=1:numOfLines
  out=matrix(NA,numOfLines+1,2,
             dimnames=list(c('SOLline1',paste('EOLline',lineIND,'SOLline',lineIND+1,sep='')[-numOfLines],paste('EOLline',numOfLines,sep='')),
                           c('lon','lat')))
  out[1,]=c(startLon,startLat)
  for(i in 1:numOfLines)
    out[i+1,]=destPoint(p=out[i,], b=angs[ifelse(i%%2==1,1,2)], d=lineLengthm)
  
  spout=SpatialPoints(out,proj4string=proj4string)
  rot=coordinates(elide(obj=spout,rotate=rotationdeg,center=coordinates(spout)[1,]))
  colnames(rot)=c('lon','lat')
  if(unrotated)
    return(list(unrotatedGeogs=out,rotatedGeogs=rot))
  return(rot)
}

#'Centre an zig-zag line transect survey on a given position
#'
#'Centres a zig-zag survey on a desired latitude and longitude
#'@param centreLon Desired centre location of survey
#'@param centreLat Desired centre location of survey
#'@param proj4string projection string of class \link{CRS-class}
#'@param tolerance maximum distance (in metres) between desired survey centre and realised survey centre
#'@param ... other arguments to be passed into \\link{zigzagSurvey}
#'@return Line transect coordinates as specified in \link{zigzagSurvey}
#'@details The call of \link{zigzagSurvey} has \code{unrotated=FALSE}
#'@export
#'@examples
#'\dontrun{
#'coords=centreZigZagOnPosition(centreLon=-33,centreLat=-57,lineLengthkm=60,startBearingdeg=30,
#'rotationdeg=10,numOfLines=21)
#'plot(0,0,xlim=range(coords[,1]),ylim=range(coords[,2]),
#'type='n',xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'points(-100,-60,col='purple',pch=19,cex=2)
#'points(geomean(coords),col='red',pch=19,cex=1)
#'legend('topright',c('Beginning','End','Desired centre','Actual centre'),
#'  col=c('blue','blue','purple','red'),pch=c(17,15,19,19),pt.cex=c(1,1,2,1))
#'  }
centreZigZagOnPosition=function(centreLon,centreLat,
                                proj4string=CRS("+proj=longlat +datum=WGS84"),tolerance=20,...)
{
  coords=zigzagSurvey(startLon=centreLon,startLat=centreLat,proj4string=proj4string,
                      unrotated=FALSE,...)
  centreCoords=(geomean(coords))
  shiftV=c(centreLon-centreCoords[1],centreLat-centreCoords[2])
  coords=coordinates(elide(obj=SpatialPoints(coords,proj4string=proj4string),shift=shiftV))
  colnames(coords)=c('lon','lat')
  shiftedCentreCoords=geomean(coords)
  errorD=distHaversine(p1=c(centreLon,centreLat), p2=geomean(coords))
  if(errorD>tolerance)
    stop(tolerance,' m tolerence between desired and realised survey centre exceeded.')
  message('Difference between desired and realised survey centre = ', round(errorD,2))
  return(coords)
}

#'Centre a regular rectangular survey on a given position
#'
#'Centres a regular rectangular survey on a desired latitude and longitude
#'@param centreLon Desired centre location of survey
#'@param centreLat Desired centre location of survey
#'@param proj4string projection string of class \link{CRS-class}.  If NULL defaults to "+proj=longlat +datum=WGS84"
#'@param tolerance maximum distance (in metres) between desired survey centre and realised survey centre
#'@param ... other arguments to be passed into \link{lawnSurvey}
#'@return Line transect coordinates (lon, lat) as specified in \link{lawnSurvey}
#'@export
#'@examples
#'\dontrun{
#'coords=centreLawnOnPosition(centreLon=-170,centreLat=-60,lineLengthkm=2,lineSpacingkm=0.5,
#'startBearingdeg=30,numOfLines=5)
#'plot(0,0,xlim=range(coords[,1]),ylim=range(coords[,2]),type='n',
#'xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'points(-170,-60,col='purple',pch=19,cex=2)
#'points(geomean(coords),col='red',pch=19,cex=1)
#'legend('bottomright',c('Beginning','End','Desired centre','Actual centre'),
#'  col=c('blue','blue','purple','red'),pch=c(17,15,19,19),pt.cex=c(1,1,2,1))
#'}
centreLawnOnPosition=function(centreLon,centreLat,
                              proj4string=NULL,tolerance=20,...)
{
  
  if(is.null(proj4string))  proj4string=CRS("+proj=longlat +datum=WGS84")
  coords=lawnSurvey(startLon=centreLon,startLat=centreLat,...)
  centreCoords=(geosphere::geomean(coords))
  shiftV=c(centreLon-centreCoords[1],centreLat-centreCoords[2])
  coords=coordinates(maptools::elide(obj=SpatialPoints(coords,proj4string=proj4string),shift=shiftV))
  colnames(coords)=c('lon','lat')
  shiftedCentreCoords=geosphere::geomean(coords)
  errorD=geosphere::distHaversine(p1=c(centreLon,centreLat), p2=geosphere::geomean(coords))
  if(errorD>tolerance)
    stop(tolerance,' m tolerence between desired and realised survey centre exceeded.')
  message('Difference between desired and realised survey centre = ', round(errorD,2),' m')
  return(coords)
}


#'Write a map info file for import into Echoview.
#'
#'This function writes a polygon MIF file for import into Echoview and will
#'typically be used to export survey line transects from, for example, 
#'\link{centreLawnOnPosition}.
#'@param coords a set of coordinates as a two column (longitude and latitude) matrix.
#'@param pathAndFileName character string MIF file export path and filename
#'@param pointNameExport Boolean (default=FALSE) export the name of each point.  
#'@param pointNameScaleFactor =c(0.005,0.03) see details
#'@details Exporting the name of each point uses the row name in the \code{coords} argument.  
#'
#'The \code{pointNameScaleFactor} argument is a two element numeric vector specifying the bounding 
#'box for each point name.  The bounding box is calculated for point name i by specifying the lower left hand corner 
#'of the point name position, with the upper right hand corner specified as 
#'c(lon_i,lat_i)+diff(range(lon_1..n,lat_1..n))*pointNameScaleFactor where n is the total number of points i.e nrow(coords)
#'@return Nothing
#'@examples
#'\dontrun{
#'coords=centreLawnOnPosition(centreLon=-33,centreLat=-55.5,lineLengthkm=100,lineSpacingkm=10,
#'startBearingdeg=30,numOfLines=50)
#'exportMIF(coords=coords,pathAndFileName='c:\\Users\\martin_cox\\Documents\\test4.mif')
#'}
exportMIF=function(coords,pathAndFileName,pointNameExport=FALSE,pointNameScaleFactor=c(0.05,0.03))
{
  textV=c('VERSION 300',
          'COLUMNS 0',
          'DATA',
          'PLINE',nrow(coords))
  for(i in 1:length(textV)) write.table(textV[i],pathAndFileName,quote=F,
                                        append=ifelse(i==1,F,T),row.names=F,
                                        col.names=F)
  
  if(pointNameExport){
    if(length(rownames(coords))!=0){
      dgeog=apply(coords,2,function(x)diff(range(x)))*pointNameScaleFactor
      for(i in 1:nrow(coords)){
        textV=paste('TEXT','\"',rownames(coords)[i],'\"')
        write.table(textV,pathAndFileName,quote=F,
                    append=T,row.names=F,col.names=F,sep='\t')
        write.table(cbind(coords[i,1],coords[i,2],coords[i,1]+dgeog[1],coords[i,2]+dgeog[1]),
                    pathAndFileName,quote=F,
                    append=T,row.names=F,col.names=F,sep='\t')
      }
    }
    else {stop('No row names found in coords ARG. Set pointNameExport=FALSE or add row names!')}
  }  
  write.table(coords,pathAndFileName,quote=F,
              append=T,row.names=F,col.names=F,sep='\t')
  return(paste('')) 
}


#' Export integration by regions by cells from an Echoview acoustic variable
#' 
#' This function performs integration by regions by cells for a specified region class and exports the results using COM scripting.
#' @param EVFile An Echoview file object
#' @param acoVarName A string containing the name of an Echoview acoustic variable
#' @param regionClassName A string containing the name of an Echoview region class
#' @param exportFn export filename and path 
#' @param dataThreshold An optional data threshold for export
#' @return a list object with one element, $msg: message for processing log.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVIntegrationByRegionsByCellsExport(EVFile, "120 aggregations", "aggregations", exportFn = "~~/KAOS/EVIntegrationByRegionsExport_example.csv")
#' }
EVIntegrationByRegionsByCellsExport <- function (EVFile, acoVarName, regionClassName, exportFn,
                                                 dataThreshold = NULL) {
  
  acoVarObj <- EVAcoVarNameFinder(EVFile = EVFile, acoVarName = acoVarName)
  msgV      <- acoVarObj$msg
  acoVarObj <- acoVarObj$EVVar
  EVRC      <- EVRegionClassFinder(EVFile = EVFile, regionClassName = regionClassName)
  msgV      <- c(msgV, EVRC$msg)
  RC        <- EVRC$regionClass
  if (is.null(dataThreshold)) {
    msg <- paste(Sys.time(),' : Removing minimum data threshold from ', acoVarName, sep = '')
    message(msg)
    msgV   <- c(msgV,msg)
    varDat <- acoVarObj[["Properties"]][["Data"]]
    varDat[['ApplyMinimumThreshold']] <- FALSE
  } else {
    msg <- EVminThresholdSet(varObj=acoVarObj,thres= dataThreshold)$msg
    message(msg)
    msgV <- c(msgV, msg)
  }
  
  msg <- paste(Sys.time(), ' : Starting integration and export of ', regionClassName, sep = '')
  message(msg)
  success <- acoVarObj$ExportIntegrationByRegionsByCells(exportFn, RC)
  
  if (success) {
    msg <- paste(Sys.time(), ' : Successful integration and export of ', regionClassName, sep = '')
    message(msg)
    msgV <- c(msgV, msg)
  } else {
    msg <- paste(Sys.time(), ' : Failed to integrate and/or export ', regionClassName, sep = '')
    warning(msg)
    msgV <- c(msgV, msg)
  } 
  
  invisible(list(msg = msgV))
}

#' Export underlying data for an acoustic variable
#'
#' This function exports underlying data for an Echoview acoustic variable using COM scripting
#' @param EVFile An Echoview file COM object
#' @param variableName a string containing the name of an EV acoustic variable
#' @param filePath a string containing the file path and name to save the exported data to
#' @param pingRange = c(-1,-1) ping range to export 
#' @return a list object with 1 element: message for progessing log
#' @keywords Echoview COM scripting
#' @details pingRange defaults are all pings i.e. start=-1 and stop=-1
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj = COMCreate('EchoviewCom.EvApplication')
#' EVFile = EVOpenFile(EVAppObj, '~~KAOS/KAOStemplate.EV')$EVFile
#' EVExportUnderlying(EVFile = EVFile, variableName = '38 seabed and surface excluded', pingRange = c(1, 100), filePath = '~~Desktop/test.csv')
#'}
EVExportUnderlying <- function (EVFile, variableName, pingRange = c(-1, -1), filePath) {
  
  acoustic.var <- EVAcoVarNameFinder(EVFile, variableName)$EVVar
  
  #check that the acoustic variable exists
  if (is.null(acoustic.var)) {
    msg <- paste(Sys.time(), "Error:", variableName, "is not an acoustic variable", sep = " ")
    warning(msg)
  } else {
    
    export.data <- acoustic.var$ExportData(filePath, pingRange[1], pingRange[2])
    
    if (export.data) {
      msg <- paste(Sys.time(), 'Success: Exported integration by cells for variable', variableName, sep = " ")
      message(msg)
    } else {
      msg <- paste(Sys.time(), 'Error: Failed to export data')
      warning(msg)
    }
  }
  invisible(list(msg = msg))
}

#' Create an EV Editable Line from an existing line
#'
#' This function creates  an EV Line in an EV file object by name using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param lineNameToCopy a string containing the name of the line to copy
#' @param editableLineName =NULL desired editable line name.If null, EDIT is appended to lineNameToCopy and used as editable line name.
#' @param Multiply =1 multiply source line depth when creating new editable line
#' @param Add =0 add to source line depth when creating new editable line
#' @param SpanGaps =FALSE span gaps when creating new editable line
#' @return an Echoview editable line object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVFindLineByName}} \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj, '~~KAOS/KAOStemplate.EV')$EVFile
#' 
#' EVCreateEditableLine(EVFile = EVFile, 
#'     lineNameToCopy='Offset seabed',
#'     editableLineName='Editable Seabed line')
#'}
EVCreateEditableLine <- function (EVFile, lineNameToCopy, editableLineName = NULL, 
                                  Multiply = 1, Add = 0, SpanGaps = FALSE) {
  
  
  lineToCopy <- EVFindLineByName(EVFile=EVFile,lineName=lineNameToCopy)
  prenbrLines <- EVFile[["Lines"]]$Count()
  
  newEditLine <- EVFile[["Lines"]]$CreateOffsetLinear(lineToCopy, Multiply, Add,SpanGaps)
  
  if (EVFile[["Lines"]]$Count() == prenbrLines) {
    stop('Failed to create editable line')
  }
  if (is.null(editableLineName)) {
    editableLineName = paste('EDIT', lineNameToCopy, sep = '')
  }
  
  renameFlag <- EVRenameLine(EVFile, evLine = newEditLine, newName = editableLineName)
  
  if (!renameFlag) {
    msg <- paste(Sys.time(), ' Rename failed, line name already exisits, attempting to force line over write....')
    message(msg)
    lineToOverwrite=EVFindLineByName(EVFile=EVFile,lineName=editableLineName)
    msg <- paste(Sys.time(),'Over writing line name ',lineToOverwrite$Name())
    message(msg)
    lineToOverwrite$OverwriteWith(newEditLine)
    msg <- paste(Sys.time(),' Deleting ',newEditLine$Name(),'...')
    EVFile[['Lines']]$Delete(newEditLine)
  } else {
    msg <- "Finished creating new line"
    lineToOverwrite = newEditLine
  }
  
  invisible(list(msg = msg, newEditLineObj = lineToOverwrite))
}

#' Export an Echoview line file for an existing line associated with an Acoustic Variable.
#'
#' This function exports an EV line definition file (.evl) for a line associated with an acoustic variable.
#' @param EVFile An Echoview file COM object
#' @param acoVar Acoustic variable name that the line to be exported is associated with.
#' @param lineNameToExport a string containing the name of the line to export
#' @param pathAndFileName Path and filename of evl file.
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVFindLineByName}} \code{\link{EVCreateEditableLine}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj, '~\\example1.EV')$EVFile
#' 
#' EVLineExportFromAcousticVarEVL(EVFile=EVFile,
#'    acoVar='120 seabed and surface excluded',
#'    lineNameToExport='Fixed depth 250 m',
#'    pathAndFileName=  '~~KAOS/test2.evl')
#'}
#'
EVLineExportFromAcousticVarEVL <- function(EVFile, acoVar, lineNameToExport, pathAndFileName){
  
  msgV <- paste(Sys.time(), 'Finding line name ', lineNameToExport, 'to export')
  message(msgV)
  
  lineToExport <- EVFindLineByName(EVFile = EVFile, lineName = lineNameToExport)
  
  msg <- paste(Sys.time(),'Writing line to ', pathAndFileName)
  message(msg)
  
  msgV <- c(msgV,msg)
  
  var <- EVAcoVarNameFinder(EVFile = EVFile, acoVar)$EVVar
  var$ExportLine(lineToExport, pathAndFileName, -1, -1)
  
  msg <- paste(Sys.time(), 'Line exported to evl file')  
  message(msg)
  msgV <- c(msgV, msg)
  
  invisible(list(msg = msgV))
}

#' Set the minimum and maximum display depth for an acoustic variable.
#'
#' This function changes the maximum and minimum depth displayed in the Echogram for an acoustic variable using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param acoVar Acoustic variable object to change display threshold
#' @param minDepth The minimum depth to display in metres
#' @param maxDepth The maximum depth to display in meters
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVAcoVarNameFinder}} 
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj, 'KAOS/KAOStemplate.EV')$EVFile
#' acoVar <- EVAcoVarNameFinder(EVFile, acoVarName = "38 seabed and surface excluded")$EVVar
#' EVSetAcoVarDisplayDepth(EVFile, acoVar, 5, 250)
#'}
#'
EVSetAcoVarDisplayDepth <- function(EVFile, acoVar, minDepth, maxDepth) {
  
  if (class(minDepth) != "numeric" | class(maxDepth) != "numeric") {
    stop("Error: Input numeric depths")
  }
  
  try({
    acoVar[["Properties"]][["Display"]][["LowerLimit"]] <- maxDepth
  }, silent = TRUE) 
  
  try({
    acoVar[["Properties"]][["Display"]][["UpperLimit"]] <- minDepth
  }, silent = TRUE)   
  
  
  new_maximum <- acoVar[["Properties"]][["Display"]][["LowerLimit"]]
  new_minimum <- acoVar[["Properties"]][["Display"]][["UpperLimit"]]
  
  if (new_maximum == maxDepth) {
    message(paste("Success: Changed the maximum depth to", maxDepth, "m"))
  } else {
    warning("Error: Could not change the maximum depth")
  }
  
  if (new_minimum == minDepth) {
    message(paste("Success: Changed the minimum depth to", minDepth, "m"))
  } else {
    warning("Error: Could not change the minimum depth")
  }
  
}

#' Export integration by region name by cells for an acoustic variable
#'
#' This function exports the integration by region by cells for an acoustic variable using COM scripting. Unlike EVIntegrationByRegionsByCellsExport, this function exports integration data for only for a single region name, not a region class. Note: This function will only work if the acoustic variable has a grid.
#' @param EVFile An Echoview file COM object
#' @param variableName a string containing the name of an EV acoustic variable
#' @param regionName a string containing the name of the region to export. Note: Input a single region's name, not region a class. To export by class, see EVIntegrationByRegionsByCellsExport.
#' @param filePath a string containing the file path and name to save the exported data to 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVExportIntegrationByRegionByCells(EVFile = EVFile, 
#' variableName = '38 seabed and surface excluded', 'Region1', 
#' filePath = '~~/KAOS/EVExportIntegrationByRegionByCells_example.csv')
#'}
EVExportIntegrationByRegionByCells <- function (EVFile, variableName, regionName, filePath) {
  
  acoustic.var <- EVAcoVarNameFinder(EVFile = EVFile, variableName)$EVVar
  
  EVRegion <- EVFindRegionByName(EVFile, regionName)
  
  export.data <- acoustic.var$ExportIntegrationBySingleRegionByCells(filePath, EVRegion)
  
  if (export.data) {
    message(paste("Success: Exported data for region", regionName, "in variable", variableName))
  } else {
    warning(paste("Data not exported"))
  }
  
}