#' Create an Echoview EVR (Region definition) file from depth and time intervals 
#' 
#' This function generates Echoview Region definitions according to predefined starting and end depts at given times.
#' The results are save as .EVR file for import into Echoview
#' 
#' @param depthStart = numeric Start depth of the new region (can be multiple)
#' @param depthStop = numeric End depth of the new region (can be multiple)
#' @param dateStart = character Starting date as %Y%m%d
#' @param dateStop = character End date as %Y%m%d
#' @param timeStart = character Starting time %H%M%S
#' @param timeStop = character End time %H%M%S
#' @param rName = character Name of the region
#' @param rClass = character Define class, otherwise it will be Unclassified
#' @param rType = Region Type - numeric 0-3 2 - Marker
#'                                   1 - Analysis
#'                                   0 - Bad Data (no data)
#'                                   4 - Bad Data (empty water)
#' @param dir = character Path to the output folder
#' @param fn = character Name of the output filename
#' @paramrNotes = character Region Notes
#' @return saves an Echoview EVR (Region definition) files
#' @keywords Echoview com scripting
#' @seealso 
#' @examples
#' \dontrun{
#'  }

makeEVRfromDT<-function(depthStart, # in meter from surface
                        depthStop,# in meter from surface
                        dateStart,# %Y%m%d - character
                        dateStop,# %Y%m%d - character
                        timeStart, # %H%M%S - character
                        timeStop, # %H%M%S - character
                        rName = "Region",
                        rClass = "Selection",
                        rType = 1,
                        dir=NULL, #set folder for raw data
                        fn=NULL, #raw file name if known
                        rNotes="") #character provideing the name for the Echoview Region
{   
  if(!all(c(length(depthStart),length(depthStop),length(timeStart))==length(timeStop)))
    stop('ARGS depths and times must have the same length.')
  opFile=paste(fn,'.evr',sep='')
  if(!is.null(dir))
    opFile=paste(dir,opFile,sep='')
  
  dateStart=as.character(dateStart)
  dateStop=as.character(dateStop)
  timeStart= substr(paste0(as.character(timeStart),'0000'),1,10)
  timeStop=substr(paste0(as.character(timeStop),'0000'),1,10)
  
  depthStart=formatC(depthStart,digits=6,format='f')
  depthStop=formatC(depthStop,digits=6,format='f')
  
  cat('EVRG 7 4.85.13.16055',file=opFile,append=FALSE, sep='\n')
  cat(length(depthStart),file=opFile, append=TRUE, sep ='\n\n')
  
  for(i in 1:length(depthStart)){
    message('Creating region for ',stationID,'.  Start depth = ',round(as.numeric(depthStart[i]),2),
            ' m. Stop depth = ',round(as.numeric(depthStop[i]),2), 'm')
    zRng <- range(depthStart[i],depthStop[i])
    cat(paste(13,4,i,0,3,-1,1,dateStart[i],timeStart[i],zRng[1],
              dateStop[i],timeStop[i],zRng[2]),file=opFile,  append=TRUE, sep ='\n')
    cat("1",append=TRUE, file=opFile,'\n')
    cat(rNotes,append=TRUE, file=opFile, '\n')
    cat("0",file=opFile, append=TRUE, sep ='\n')
    cat(rClass, append=TRUE, file=opFile,"\n")
    cat(paste(dateStart[i],timeStart[i],depthStart[i],dateStart[i],timeStart[i],depthStop[i],
              dateStop[i],timeStop[i],depthStop[i],dateStop[i],timeStop[i],depthStart[i],rType,collapse=' '), 
        append=TRUE, file=opFile,sep ='\n')
    cat(paste0(rName,i),append=TRUE, file=opFile, sep='\n')
  }
  message('Echoview region (.evr) file written to: ',file=opFile)
}  
