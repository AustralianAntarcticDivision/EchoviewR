#' Change the single target settings for Echoview single target virtual variable
#' 
#'Change the single target settings for Echoview virtual variable 
#'  
#'@import RDCOMClient
#'@param EVFile ("COMIDispatch) An Echoview file COM object 
#'@param acoVarName (character) Name of Echoview single target virtual variable
#'@param ExcludeAboveLine (character) Shallower exclusion line
#'@param ExcludeBelowLine (character) Deeper exclusion line
#'@param MinimumPulseLength (numeric) Minimum pulse length
#'@param MaximumPulseLength (numeric) Maximum pulse length
#'@param PLDL (numeric) Pulse length determination level
#'@param TsThreshold (numeric) minimum target strength threshold (dB)
#'@param verbose = FALSE (boolean) 
#'@details This function will work with single target detection viritual variables 
#'@note THis function has only been tested using Simrad EK60 split-beam data
#'@return list $settings 2 column data frame of pre-change post-change parameter values, $msg= vector of messages
#'@export
EVsingleTargetPars=function(EVFile,acoVarName,
                            ExcludeAboveLine=NULL,
                            ExcludeBelowLine=NULL,
                            MinimumPulseLength=NULL,
                            MaximumPulseLength=NULL,
                            PLDL=NULL,
                            TsThreshold=NULL,
                            verbose=FALSE){
  #find virtual varialble
  finder=EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=finder$msg #we will keep track of messages produced by the function here
  acoVar=finder$EVVar #this is the COM address of the acoustic virtual variable of interest
  #check acoustic variable type
  varType=acoVar$Operator()
  if(!(varType %in% c(66,74,75,114,115,124))){
    msg=paste(Sys.time(),' : input variable ',acoVarName,' is not a single-target-type virtual acoustic variable')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(settings=NULL,msg=msgV))
  }
  #get current settings
  
  STparametersDataFrame=data.frame(matrix(NA,2,6))
  nameV=c('ExcludeAboveLine','ExcludeBelowLine','MinimumPulseLength',
                       'MaximumPulseLength','PLDL','TsThreshold')
  names(STparametersDataFrame)=nameV
  
  row.names(STparametersDataFrame)=c('pre','post')
  #Single target detection parameters settings:
  singleTargetPars=acoVar[['Properties']][['SingleTargetDetectionParameters']]
  
  STparametersDataFrame$ExcludeAboveLine[1]=singleTargetPars$ExcludeAboveLine()
  STparametersDataFrame$ExcludeBelowLine[1]=singleTargetPars$ExcludeBelowLine()
  
  STparametersDataFrame$MinimumPulseLength[1]=singleTargetPars$MinimumPulseLength()
  STparametersDataFrame$MaximumPulseLength[1]=singleTargetPars$MaximumPulseLength()
  
  STparametersDataFrame$PLDL[1]=singleTargetPars$PLDL()
  STparametersDataFrame$TsThreshold[1]=singleTargetPars$TsThreshold()
  
  if(verbose){
  message('---------- PRE-CHNAGE SINGLE TARGET PARAMETERS -----------------------')  
  message('Exclude above line = ', STparametersDataFrame$ExcludeAboveLine[1])
  message('Exclude below line = ', STparametersDataFrame$ExcludeBelowLine[1])
  message('Minimum pulse length = ', STparametersDataFrame$MinimumPulseLength[1])
  message('Maximum pulse length = ', STparametersDataFrame$MaximumPulseLength[1])
  message('Pulse length determination level (PLDL) = ', STparametersDataFrame$PLDL[1],' dB')
  message('Target strength threshold = ', STparametersDataFrame$TsThreshold[1],' dB')
  message('------------------------------------------------------------------------')  }
  
  #Change settings:
  msg=paste(Sys.time(),' : Changing single target detection settings for ',acoVarName)
  message(msg)
  msgV=c(msgV,msg)
  
  if(!is.null(ExcludeAboveLine)) singleTargetPars[['ExcludeAboveLine']]=ExcludeAboveLine
  if(!is.null(ExcludeBelowLine))singleTargetPars[['ExcludeBelowLine']]=ExcludeBelowLine
  if(!is.null(MinimumPulseLength)) singleTargetPars[['MinimumPulseLength']]=MinimumPulseLength
  if(!is.null(MaximumPulseLength)) singleTargetPars[['MaximumPulseLength']]=MaximumPulseLength
  if(!is.null(PLDL)) singleTargetPars[['PLDL']]=PLDL
  if(!is.null(TsThreshold)) singleTargetPars[['TsThreshold']]=TsThreshold
  
  #change settings were changed:
  singleTargetPars=acoVar[['Properties']][['SingleTargetDetectionParameters']]
  STparametersDataFrame$ExcludeAboveLine[2]=singleTargetPars$ExcludeAboveLine()
  STparametersDataFrame$ExcludeBelowLine[2]=singleTargetPars$ExcludeBelowLine()
  
  STparametersDataFrame$MinimumPulseLength[2]=singleTargetPars$MinimumPulseLength()
  STparametersDataFrame$MaximumPulseLength[2]=singleTargetPars$MaximumPulseLength()
  
  STparametersDataFrame$PLDL[2]=singleTargetPars$PLDL()
  STparametersDataFrame$TsThreshold[2]=singleTargetPars$TsThreshold()
  
  if(verbose){
    message('----------- POST-CHANGE SINGLE TARGET PARAMETERS -----------------------')  
    message('Exclude above line = ', STparametersDataFrame$ExcludeAboveLine[2])
    message('Exclude below line = ', STparametersDataFrame$ExcludeBelowLine[2])
    message('Minimum pulse length = ', STparametersDataFrame$MinimumPulseLength[2])
    message('Maximum pulse length = ', STparametersDataFrame$MaximumPulseLength[2])
    message('Pulse length determination level (PLDL) = ', STparametersDataFrame$PLDL[2],' dB')
    message('Target strength threshold = ', STparametersDataFrame$TsThreshold[2],' dB')
    message('------------------------------------------------------------------------')  }
   
  STparametersDataFrame=data.frame(t(STparametersDataFrame))
  print(STparametersDataFrame)
  STparametersDataFrame<<-  STparametersDataFrame
  row.names(STparametersDataFrame)=nameV
  names(STparametersDataFrame)=c('pre','post')

    return(list(settings=STparametersDataFrame,msg=msgV))
}

