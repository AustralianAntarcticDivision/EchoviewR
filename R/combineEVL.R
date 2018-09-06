#' Combine Echoview line objects
#' 
#' This function combines evl class objects into a single .evl file
#' @param evlL list of objects of class evl, typically resulting from a class of readevl2R
#' @return an evl class object
#' @seealso \code{\link{readevl2R}} \code{\link{writeRevl2file}}
#' @export
combineEVL=function(evlL){
  message(Sys.time(),' : attempting to combine EVL objects.')
  classV=sapply(evlL,class)
  if(any(classV!='evl')){
    warning(Sys.time(),' : one or more objects in the evlL ARG are not of class evl')
    return(NULL)
  }
  timeOrder=order(sapply(evlL,function(x) min(x$timeStamp)))
  evlL=evlL[timeOrder]
  bodyEVLout=do.call('rbind',lapply(evlL,function(x) x$bodyEVL))
  EVLout=evlL[[1]]
  EVLout$n=nrow(bodyEVLout)
  EVLout$timing[2]=evlL[[length(evlL)]]$timing[2]
  EVLout$bodyEVL=bodyEVLout
  EVLout$timeStamp=sapply(evlL,function(x) as.character(x$timeStamp))
  EVLout$msg=sapply(evlL,function(x) as.character(x$msg))
  EVLout$msg=c(EVLout$msg,paste(Sys.time(),': this line was created by combining EVL files'))
  return(EVLout)
}
