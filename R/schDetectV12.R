#' Change schools detection settings for Echvoiew >=V12 
#' 
#' This function changes schools detection settings in Echoview version >=12
#' @param EVApp = EV application COM object
#' @param EVFile =NULL not used An Echoview file COM object
#' @param varObj =NULL not used the EV acoustic object to change schools detection parameters for
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
#' @seealso \code{\link{EVSchoolsDetSet}}
#' @examples 
#' \dontrun{
#' library(EchoviewR)
#'EVApp=NULL
#'EVApp=EchoviewR::StartEchoview()
#'EVFile=EchoviewR::EVOpenFile(EVAppObj = EVApp,
#'                             fileName='x:\\KAOS\\20120519_KAOS_all_2.EV')$EVFile
#'
#'EVSchoolsDetSetV12(EVApp=EVApp,
#'distanceMode="GPSDistance",
#'maximumHorizontalLink=20,
#'maximumVerticalLink=10,
#'minimumCandidateHeight=10,
#'minimumCandidateLength=40,
#'minimumSchoolHeight=10,
#'minimumSchoolLength=50)
#' 
#' }
EVSchoolsDetSetV12  <- function (EVApp,EVFile=NULL, varObj=NULL, distanceMode,
                              maximumHorizontalLink,
                              maximumVerticalLink,
                              minimumCandidateHeight,
                              minimumCandidateLength,
                              minimumSchoolHeight,
                              minimumSchoolLength) {
  setVec <- c(maximumHorizontalLink, maximumVerticalLink, minimumCandidateHeight, 
              minimumCandidateLength, minimumSchoolHeight,minimumSchoolLength)
  f_strp=function(x,num=TRUE) {out=strsplit(x,'\\|')[[1]][2]
  out=gsub(' ','',out)
  if(num) out=as.numeric(out)
  return(out)}
  f_setr = function(EVApp, varName, newVarVal
  )
  {
    res=EVApp$Exec(paste(varName,'= |',newVarVal))
    if(length(grep('ERROR',res))>0) stop(res)
    invisible(EVApp$Exec(varName))
  }
  if (!all(is.numeric(setVec))) {
    stop('Non-numeric ARG in school detection distance settings')
  }
  
  #get current school detection parameters
  pre.distmode     <- f_strp(EVApp$Exec('SchoolsDistanceMode'),num=F)
  pre.maxhzlink    <- f_strp(EVApp$Exec('SchoolsMaximumHorizontalLinkDistance'))
  pre.maxvtlink    <- f_strp(EVApp$Exec('SchoolsMaximumVerticalLinkDistance'))
  pre.mincandHt    <- f_strp(EVApp$Exec('SchoolsMinimumCandidateHeight'))
  pre.mincandLen   <- f_strp(EVApp$Exec('SchoolsMinimumCandidateLength'))
  pre.minSchoolHt  <- f_strp(EVApp$Exec('SchoolsMinimumTotalHeight'))
  pre.minSchoolLen <- f_strp(EVApp$Exec('SchoolsMinimumTotalLength'))
  preSettingDistances <- c(pre.maxhzlink = pre.maxhzlink,
                           pre.maxvtlink = pre.maxvtlink, pre.mincandHt = pre.mincandHt, 
                           pre.mincandLen = pre.mincandLen, pre.minSchoolHt = pre.minSchoolHt, 
                           pre.minSchoolLen = pre.minSchoolLen)
  
  #set school parameters
  f_setr(EVApp=EVApp,varName='SchoolsDistanceMode',distanceMode)
  f_setr(EVApp=EVApp,varName='SchoolsMaximumHorizontalLinkDistance',maximumHorizontalLink)
  f_setr(EVApp=EVApp,varName='SchoolsMaximumVerticalLinkDistance',maximumVerticalLink)
  f_setr(EVApp=EVApp,varName='SchoolsMinimumCandidateHeight',minimumCandidateHeight)
  f_setr(EVApp=EVApp,varName='SchoolsMinimumCandidateLength',minimumCandidateLength)
  f_setr(EVApp=EVApp,varName='SchoolsMinimumTotalHeight',minimumSchoolHeight)
  f_setr(EVApp=EVApp,varName='SchoolsMinimumTotalLength',minimumSchoolLength)
  

  #check settings have been applied by getting current (post change) school detection parameters
  post.distmode     <- f_strp(EVApp$Exec('SchoolsDistanceMode'),num=F)
  post.maxhzlink    <- f_strp(EVApp$Exec('SchoolsMaximumHorizontalLinkDistance'))
  post.maxvtlink    <- f_strp(EVApp$Exec('SchoolsMaximumVerticalLinkDistance'))
  post.mincandHt    <- f_strp(EVApp$Exec('SchoolsMinimumCandidateHeight'))
  post.mincandLen   <- f_strp(EVApp$Exec('SchoolsMinimumCandidateLength'))
  post.minSchoolHt  <- f_strp(EVApp$Exec('SchoolsMinimumTotalHeight'))
  post.minSchoolLen <- f_strp(EVApp$Exec('SchoolsMinimumTotalLength'))
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
