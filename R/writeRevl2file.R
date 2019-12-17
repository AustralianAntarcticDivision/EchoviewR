#' Write an Echoview line (.evl) file 
#' @param evlObj An R object of class 'evl' resulting from a call of readevl2R()
#' @param pathFn path and filename of evl file to be create
#' @return list object: [1]$success; [2] msg message log  
#' @seealso \code{\link{readevl2R}}
#' @export
writeRevl2file=function(evlObj,pathFn)
{
  msgV=paste(Sys.time(),': Creating EVL file:',pathFn)
  message(msgV)
  out=list(success=NULL,msg=msgV)
  if(class(evlObj)!='evl')
  {
    msg=paste(Sys.time(),': failed to create EVL file ',pathFn)
    msgV=c(msgV,msg)
    warning(msg)
    out$msg=msgV
    out$success=FALSE
    return(out)
  }
  write.table(evlObj$line1,pathFn,sep=' ',
              col.names = F,row.names = F,quote=F)
  write.table(evlObj$n,pathFn,sep=' ',append=TRUE,
              col.names = F,row.names = F,quote=F)
  write.table(evlObj$bodyEVL,pathFn,sep=' ',append=TRUE,
              col.names = F,row.names = F,quote=F)
  msg=paste(Sys.time(),': created EVL file',pathFn)
  message(msg)
  msgV=c(msgV,msg)
  out$flag=TRUE
  out$msg=msgV
  return(out)
}  
