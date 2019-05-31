#' Control an Echoview data generator variable using COM scripting.
#' 
#' This function modifies the data generator settings in an existing Echoview data generator virtual variable. 
#' @param EVFile An Echoview file COM object
#' @param acoVarName Acoustic variable name 
#' @param Algorithm =c("Constant","Tvg","ConstantTvg") Data generator algorithm type
#' @param OutputType =c("Sv","TS","Other dB","Linear") output type
#' @param constantValue = NULL (numerical) constant value for the data generator.
#' @param ApplyTvgOverride = NULL (boolean) Appply TVG override
#' @param TvgOverride = NULL (numerical) TVG override value.
#' @return a list object with two elements.  $success: ping subset successfully changed, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples 
#' \dontrun{
#' EVAcoVarDataGenerator(EVFile=EVFile,acoVarName='Data generator 1',
#'Algorithm='Constant',
#'OutputType='Linear',
#'constantValue=0.5,
#'ApplyTvgOverride=NULL,
#'TvgOverride=NULL)
#'
#' }
EVAcoVarDataGenerator=function(EVFile,acoVarName,
                               Algorithm=c("Constant","Tvg","ConstantTvg"),
                               OutputType=c("Sv","TS","Other dB","Linear"),
                               constantValue=NULL,
                               ApplyTvgOverride=NULL,
                               TvgOverride=NULL)
{
  varObj<-EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=varObj$msg
  varObj<-varObj$EVVar
  msg=paste(Sys.time(),' : Changing data generator specification in ',varObj$Name(),sep='')
  message(msg)
  msgV=c(msgV,msg)
  genObj=varObj[['Properties']][['Generator']]
  #get current settings:
  #OutputType
  preOutputType=genObj[['OutputType']] 
  msg=paste(Sys.time(),': Pre-change output type =',preOutputType)
  message(msg)
  msgV=c(msgV,msg)
  #Algorithm:
  preAlgorithmType=genObj[['Algorithm']] 
  msg=paste(Sys.time(),': Pre-change algorithm type =',preAlgorithmType)
  message(msg)
  msgV=c(msgV,msg)
  #Constant valuable
  preConstantVal=genObj[['Constant']] 
  msg=paste(Sys.time(),': Pre-change constant value =',preConstantVal)
  message(msg)
  msgV=c(msgV,msg)
  #Apply TVG value boolean
  preTVGoverride=genObj[['ApplyTvgOverride']] 
  msg=paste(Sys.time(),': Pre-change apply TVG override (boolean) =',preTVGoverride)
  message(msg)
  msgV=c(msgV,msg)
  #Apply TVG value 
  preTVGVal=genObj[['TvgOverride']] 
  msg=paste(Sys.time(),': Pre-change TVG override value =',preTVGVal)
  message(msg)
  msgV=c(msgV,msg)
  #end of pre-values
  ##Algorithm changes:
  alg=match.arg(Algorithm)
  genObj[['Algorithm']]=alg
  postAlgType=genObj[['Algorithm']]
  msg=paste(Sys.time(),': Post-change algorithm type =',postAlgType)
  message(msg)
  msgV=c(msgV,msg)
  #Output type:
  OutputType=match.arg(OutputType)
  genObj[['OutputType']] = OutputType
  postOutputType=genObj[['OutputType']] 
  msg=paste(Sys.time(),': Post-change output type =',postOutputType)
  message(msg)
  msgV=c(msgV,msg)
  if(length(constantValue)>0){
    genObj[['Constant']] = constantValue
    postConstantVal=genObj[['Constant']] 
    msg=paste(Sys.time(),': Post-change constant value =',postConstantVal)
    message(msg)
    msgV=c(msgV,msg)
  }
  if(length(ApplyTvgOverride)>0){
    if(class(ApplyTvgOverride)=="logical"){
      genObj[['ApplyTvgOverride']] = ApplyTvgOverride
      postTVGoverride=genObj[['ApplyTvgOverride']] 
      msg=paste(Sys.time(),': Post-change apply TVG override (boolean) =',postTVGoverride)
      message(msg)
      msgV=c(msgV,msg)} else {
      msg=paste(Sys.time(),': ApplyTvgOverride ARG must be boolean.  TVG override apply not changed')
      warning(msg)
      msgV=c(msgV,msg)
      }
  } #end tvg override
  
  if(length(TvgOverride)>0){
  if(is.numeric(TvgOverride)){
    genObj[['TvgOverride']]=TvgOverride
    postTVGVal=genObj[['TvgOverride']] 
    msg=paste(Sys.time(),': Post-change TVG override value =',postTVGVal)
    message(msg)
    msgV=c(msgV,msg)} else 
    {
      msg=paste(Sys.time(),': TvgOverride ARG must be numeric.  TVG override value not changed')
      warning(msg)
      msgV=c(msgV,msg)
    }
  }
  return(list(msg=msgV,obj=varObj))
}  