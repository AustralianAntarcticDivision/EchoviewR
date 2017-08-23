#' Make a rectangular Echoview region (EVR) file.
#' 
#' This function creates a rectangular (2D) Echoview region file.  The function can add regions to an existing EVR file.
#'@param dateStart date  stamp for the start of the region (yyymmdd)
#'@param timeStart time stamp for the start of the region (hhmmss)
#'@param dateStop date stamp for the end of the region (yyymmdd)
#'@param timeStop time stamp for the end of the region (hhmmss)
#'@param minDepth =0 Minimum depth for the region
#'@param maxDepth maximum depth for the region
#'@param regionClassName region class name
#'@param regionType region Type (integer)
#'@param outPathFn output parth and filename.
#'@param append = FALSE should be region be added to the EVR file given in outPathFn
#'@return list: $msg messages from the function call $exportFLAG Boolean TRUE EVR file created or added.  FALSE otherwise
#'@details Region types supported by Echoview: 0 = bad (no data); 1 = analysis; "2" = marker; 3 = fishtracks, and 4 = bad (empty water)
#'@references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#'@export
#'@examples 
#' \dontrun{makeRectRegion(dateStart='20170815',timeStart='001525',
#'dateStop='20170815',timeStop='001738', 
#'minDepth=20,maxDepth=100,
#'outPathFn='~/test2.evr',
#'regionClassName='ctd',append=FALSE)}
makeRectRegion=function(dateStart,timeStart,dateStop,timeStop, minDepth=0,maxDepth,outPathFn,
                        regionClassName,regionType,append=FALSE)
{
  msgV=paste(Sys.time(),' : Creating rectangular region')
  message(msgV)
  regionID=0
  if(append){
    if(file.exists(outPathFn)) {msg=paste(Sys.time(),': Loading 2D region definition file =',outPathFn)
    message(msg)
    msgV=c(msgV,msg)} else {
      msgV=paste(Sys.time(),': Region file missing (',outPathFn,')')
      warning(msg)
      msgV=c(msgV,msg)
      return(list(msg=msgV))
    }
    regionID=as.numeric(read.table(outPathFn,skip=1,nrows=1))
    allEVR=readLines(outPathFn,-1)
    allEVR[2]=regionID+1
    writeLines(allEVR,outPathFn)
  }
  if(!append){
    msg=paste(Sys.time(),' : creating new 2D Echoview region file =',outPathFn)
    message(msg)
    msgV=c(msgV,msg)
    write.table('EVRG 7 8.0.86.31499',outPathFn,sep='',quote=FALSE,col.names=F,row.names=F)
    write.table('1',outPathFn,sep='',append=T,quote=FALSE,col.names=F,row.names=F)
    
  }
  topLeftPoint=paste(dateStart,' ',timeStart,'0000 ',minDepth,sep='')
  bottomRightPoint=paste(dateStop,' ',timeStop,'0000 ',maxDepth,sep='')
  bottomLeftPoint=paste(dateStart,' ',timeStart,'0000 ',maxDepth,sep='')
  topRightPoint=paste(dateStop,' ',timeStop,'0000 ',minDepth,sep='')
  dfReg=data.frame(line1='')
  dfReg$line2=paste('13 4',regionID+1,'0 3 -1 1',topLeftPoint,bottomRightPoint)
  dfReg$line3=0
  dfReg$line4=0
  dfReg$line5=regionClassName
  dfReg$line6=paste(topLeftPoint,bottomLeftPoint,bottomRightPoint,topRightPoint,regionType)
  dfReg$line7=paste('Region',regionID+1,sep='')
  for(i in 1:ncol(dfReg)) write.table(dfReg[,i],outPathFn,sep='',quote=F,append=T,col.names=F,row.names=F)
  
  invisible(list(msg=msgV,exportFLAG=ifelse(file.exists(outPathFn),TRUE,FALSE)))
}
