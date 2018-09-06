#' Read and Echoview line (.evl) file into R
#' @param pathFn path and filename of evl file
#' @param tz ='GMT' time zone see as.POSIXlt
#' @return list object of class evl: [1]$timing ($start, $end start and end time stamps); [2] $line1 line 1 - evl file first line; 
#' [3] $n number of points in evl file; [4] $bodyEVL body of evl file; [5] $timeStamp time stamp for; [6] msg message log;;
#' [7] $pathFn path and filename of Echoview line file; [8] tz time zone  
#' @export
readevl2R=function(pathFn,tz='GMT')
{
  outputList=list(timing=NULL,line1=NULL,n=NULL,bodyEVL=NULL,
                  timeStamp=NULL,msg=NULL,pathFn=pathFn,tz=tz)
  
  msgV=paste(Sys.time(),' : attempting to read Echoview line file',pathFn)
  message(msgV)
  fFlag=file.exists(pathFn)
  if(!fFlag) {msg=paste(Sys.time(),': file',pathFn,'not found')
  msgV=c(msgV,msg)
  warning(msg)
  outputList$msg=msgV
  return(outputList)}
  
  bodyEVL=NULL
  bodyEVL=read.table(pathFn,sep=' ',skip=2)
  if(class(bodyEVL)=='NULL') {
    msg=paste(Sys.time(),': failed to read data in ',pathFn)
    msgV=c(msgV,msg)
    warning(msg)
    outputList$msg=msgV
    return(outputList)
    
  }
  bodyEVL=bodyEVL[,-c(3,6)]
  outputList$bodyEVL=bodyEVL
  head1=scan(pathFn,what='character',nlines=1,quiet=TRUE)
  locStart=as.numeric(gregexpr(pattern ='E',head1[1]))
  head1[1]=substr(head1[1],locStart,nchar(head1[1]))
  outputList$line1=paste(head1,collapse = ' ')
  
  evlHead2=scan(pathFn,what='character',nlines=1,skip=1,quiet=TRUE)
  outputList$n=as.numeric(evlHead2)
  msg=paste(Sys.time(),' : success read in',outputList$n,'rows in EVL file')
  message(msg)
  msgV=c(msgV,msg)
  
  timestamp=formatC(bodyEVL[,2],flag='0',width=10,format='fg')
  timestamp=paste(substr(timestamp,1,2),':',substr(timestamp,3,4),':',
                  substr(timestamp,5,6),'.',substr(timestamp,7,10),sep='')
  timestamp=paste(bodyEVL[,1],timestamp)
  timestamp=as.POSIXct(timestamp,format="%Y%m%d %H:%M:%OS",tz=tz)
  outputList$timeStamp=timestamp
  outputList$msg=msgV
  outputList$timing=c(min(timestamp),max(timestamp))
  names(outputList$timing)=c('start','end')
  class(outputList)='evl'
  return(outputList)
}



