#' Find Raw files based on Time and Date, providing folder containing all Raw files
#' 
#' This functions generates datetime object based on the raw data filenames and selects the required raw files according to 
#' a user defined start and end time
#' @import lubridate
#' @note  Dependencies: lubridate - Requires lubridate package for easy computation of time intervals
#' @param dir = path [character]
#' @param StartDate [character]  Date under the shape of YYYYMMDD (example: 20171127)
#' @param StartTime   [character]  Time under the shape of HHMMSS (example: 085617)
#' @param EndDate [character]  Date under the shape of YYYYMMDD (example: 20171127)
#' @param EndTime [character]  Time under the shape of HHMMSS (example: 085617)
#' @param tzraw ="UTC"[character]  timezone under which the raw files were recorded e.g.UTC or Australia/Brisbane or AEST 
#' @param tzinp ="UTC"[character]  timezone of the entered start and end points e.g.UTC or Australia/Brisbane or AEST 
#' @param timebuffer = 0 [numeric in seconds]  add a buffer around the times for which files should be selected
#' @param fileSlope = FALSE [boolean] If TRUE the raw file that partly falls within the start of the time interval.
#' @return charater array with raw data filenames 
#' @keywords Echoview com scripting
#' @export
#' @examples
#' \dontrun{
#' StartDate = 20170901
#' StartTime = "080917"
#' EndDate = 20170901
#' EndTime = 172356
#' rawdir = "F:\\RawFolder\\"
#' getRaws(dir,StartDate, StartTime,EndDate, EndTime, tzraw="UTC", tzinp="UTC")
#'  }
#'  
#'  
#dir=RAWFileDir;StartDate='20160126';StartTime='080000';EndDate='20160126';EndTime='121847';fileSlope = TRUE

getRaws <- function(dir,StartDate, StartTime,EndDate, EndTime, tzraw="UTC", tzinp="UTC", timebuffer = 0,
                    fileSlope=FALSE){
  #require(lubridate)
  lsRaw <- list.files(path=dir,pattern="*.raw") # get all raw files in folder
  lsRaw <- lsRaw[substr(lsRaw,nchar(lsRaw)-3,nchar(lsRaw))==".raw"] #Excludes files such as .raw.evi
  
  lst=strsplit(lsRaw,'-')
  f1=function(x) {
    dd=x[which(substr(x,1,1)=='D')]
    dd=substr(dd,2,nchar(dd))
    tt=x[which(substr(x,1,1)=='T')]
    tt=gsub('.raw','',tt)
    tt=substr(tt,2,nchar(tt))
    paste(dd,tt)}
  
  rawpos <- as.POSIXct(sapply(lst,f1),format="%Y%m%d %H%M%S", tz=tzraw)
  
  #Get datetime object from filenames
#  spo <- regexpr(paste0("D",substr(as.character(StartDate),1,4)),lsRaw) #where the dates start in the filename
#  days <- substr(lsRaw,spo+1,spo+8)
#  times <- substr(lsRaw, spo+11,spo+17)
#  rawpos <- as.POSIXct(paste(days,times),format="%Y%m%d %H%M%S", tz=tzraw)
  
  start_int=as.POSIXct(paste(StartDate, StartTime),format="%Y%m%d %H%M%S", 
                       tz=tzinp)-timebuffer
  stop_int=as.POSIXct(paste(EndDate, EndTime),format="%Y%m%d %H%M%S", 
                      tz=tzinp)+timebuffer
  
  timeint <- interval(start_int,stop_int) #get time interval
  
  IND=which(rawpos %within% timeint)
  if(length(IND)==0) {
    warning(Sys.time(),' : No files found.')
    return(NA)
  }
  out=lsRaw[IND]
  
  if(fileSlope){
    if((IND[1]-1)<1) {warning('fileSlope ARG ignored')} else{
      flag=rawpos[IND[1]-1]< start_int & rawpos[IND[1]] > start_int
    if(flag) out=lsRaw[c((IND[1]-1),IND)]
    }
  }
  return(out)
}

