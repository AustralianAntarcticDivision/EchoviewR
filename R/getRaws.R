#' Find Raw files based on Time and Date, providing folder containing all Raw files
#' 
#' This functions generates datetime object based on the raw data filenames and selects the required raw files according to 
#' a user defined start and end time
#' 
#' @dependencies lubridate - Requires lubridate package for easy computation of time intervals
#' @param dir = path [character]
#' @param StartDate / EndDate [character] = Date under the shape of YYYYMMDD (example: 20171127)
#' @param StartTime / EndTime [character] = Time under the shape of HHMMSS (example: 085617)
#' @param tzraw [character] = timezone under which the raw files were recorded e.g.UTC or Australia/Brisbane or AEST 
#' @param tzinp [character] = timezone of the entered start and end points e.g.UTC or Australia/Brisbane or AEST 
#' @param timebuffer [numeric in seconds] = add a buffer around the times for which files should be selected
#' @return charater array with raw data filenames 
#' @keywords Echoview com scripting
#' @seealso 
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
getRaws <- function(dir,StartDate, StartTime,EndDate, EndTime, tzraw="UTC", tzinp="UTC", timebuffer = 0){
  require(lubridate)
  lsRaw <- list.files(path=rawdir,pattern="*.raw") # get all raw files in folder
  lsRaw <- lsRaw[substr(lsRaw,nchar(lsRaw)-3,nchar(lsRaw))==".raw"] #Excludes files such as .raw.evi
  #Get datetime object from filenames
  spo <- regexpr(paste0("D",substr(as.character(StartDate),1,4)),lsRaw) #where the dates start in the filename
  days <- substr(lsRaw,spo+1,spo+8)
  times <- substr(lsRaw, spo+11,spo+17)
  rawpos <- as.POSIXct(paste(days,times),format="%Y%m%d %H%M%S", tz=tzraw)
  
  timeint <- interval(as.POSIXct(paste(StartDate, StartTime),format="%Y%m%d %H%M%S", 
                                 tz=tzinp)-timebuffer,
                      as.POSIXct(paste(EndDate, EndTime),format="%Y%m%d %H%M%S", 
                                 tz=tzinp)+timebuffer) #get time interval
  lsRaw[rawpos %within% timeint] #select only files within time interval
}
