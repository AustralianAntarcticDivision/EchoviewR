#' Get calibration settings pf an Echoview acoustic variable using COM
#' 
#' This function gets the calibration values of an Echoview acousic variable using COM. 
#' @details The function can read three types of calibration values "default", "set" and "used".
#' @param EVFile ("COMIDispatch) An Echoview file COM object 
#' @param acoVarName (character) Name of Echoview single target virtual variable
#' @param pingNumber = 0 (numeric) Ping number from which to extract calibration values.
#' @param type = =c('default','set','used') (character) Calibration value type.
#' @param verbose = FALSE (boolean) Should the calibration values be displayed in the console.
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVAcoVarNameFinder}}
#' @export
EVgetCalParValues=function(EVFile,acoVarName,pingNumber=1,type=c('default','set','used'),verbose=FALSE){
  #set-up output
  msgV=paste(Sys.time(),' : looking for calibration values for acoustic variable', acoVarName)
  message(msgV)
  
  out=list(obj=NULL,
    Name=acoVarName,
    pingNumber=NULL,
    type=NULL,
    pars=NULL,
    msg=msgV)
  
  EVVar=EVAcoVarNameFinder(EVFile=EVFile, acoVarName=acoVarName)
  EVVar=EVVar$EVVar
  msgV=c(msgV,EVVar$msg)
  out$msg=msgV
  if(is.null(EVVar)) {
    msg=paste(Sys.time(),' : Acoustic variable', acoVarName ,'not found. Exiting EVgetParValues()')
    warning(msg)
    msgV=c(msgV,msg)
    out$msg=msgV
    return(out)
  }
  out$obj=EVVar
  out$pingNumber=pingNumber
  out$type=type
  
  cal=EVVar[['Properties']][['Calibration']]
  print('here!')
  if(type=='set') calVarName=strsplit(cal$GetAllSet(pingNumber),' ')[[1]]
  if(type=='used') calVarName=strsplit(cal$GetAllUsed(pingNumber),' ')[[1]]
  if(type=='default') calVarName=strsplit(cal$GetDefault(pingNumber),' ')[[1]]
  if(length(calVarName)==0)
  {
    msg=paste(Sys.time(),': no calibration parameters of the',type,'type were found in',acoVarName)
    warning(msg)
    msgV=c(msgV,msg)
    out$msg=msgV
    return(out)
  }
  print(calVarName)
  out<<-calVarName
  parVSet=vector(mode='numeric',length=length(calVarName))
  for(i in 1:length(calVarName)) parVSet[i]=cal$Get(calVarName[i],pingNumber)
  out$pars=data.frame(parName=calVarName,parval=parVSet)
  if(verbose){
  for(i in 1:length(calVarName))
    message(paste(calVarName[i],parVSet[i],sep=' = '))}
  
  return(out)
}

