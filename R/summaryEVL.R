summary.evl=function(x){
  cat(paste(rep('-',50),collapse = ''),'\n')
  cat('ECHOVIEW LINE OBJECT SUMMARY','\n')
  cat('Echoview line file = \n')
  cat(x$pathFn,'\n')  
  cat('Start time stamp =',as.character(x$timing[1]),'\n')
  cat('Stop time stamp =',as.character(x$timing[2]),'\n')
  cat('n = ',x$n,'\n')
  cat('min. depth = ',round(min(x$bodyEVL[,3]),1),' m','\n')
  cat('max. depth = ',round(max(x$bodyEVL[,3]),1),' m','\n')
  cat(paste(rep('-',50),collapse = ''),'\n')
  
}
