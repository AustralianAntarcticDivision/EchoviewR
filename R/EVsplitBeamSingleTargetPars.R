#' Change the SPLIT-BEAM single target settings for an Echoview split-beam single target virtual variable
#' 
#'Change the split-beam single target settings for Echoview virtual variable of type split beam (method 1) and split-eabm (method 2)/
#'  
#'@import RDCOMClient
#'@param EVFile ("COMIDispatch) An Echoview file COM object 
#'@param acoVarName (character) Name of Echoview single target virtual variable
#'@param BeamCompensationModel ='SimradLOBE' c('Simrad LOBE','BioSonics','HTI','Furuno','Sonic')Beam compensation model
#'@param MaximumBeamCompensation =NULL (numeric)
#'@param MaximumStdDevOfMajorAxisAngles =NULL (numeric) maximum standard deviation of Major axis angles (degrees)
#'@param MaximumStdDevOfMinorAxisAngles =NULL (numeric) maximum standard deviation of Minor axis angles (degrees)
#'@param verbose = FALSE (boolean) 
#'@details This function will work with split-beam single target detection viritual variables (methods 1 and 2).
#'@note THis function has only been tested using Simrad EK60 split-beam data
#'@return list $settings 2 column data frame of pre-change and post-change parameter values , $msg= vector of messages
#'@export
EVsplitBeamSingleTargetPars=function(EVFile,acoVarName,
                            BeamCompensationModel= 'SimradLobe',
                            MaximumBeamCompensation=NULL,
                            MaximumStdDevOfMajorAxisAngles=NULL,
                            MaximumStdDevOfMinorAxisAngles=NULL,
                            verbose=FALSE){
  #switch functions
  beamText2enum=function(txt) {switch(txt,
                                      Biosonics = 1,
                                      Furuno = 3,
                                      Hti=2,
                                      None=4,
                                      SimradLobe=0,
                                      Sonic=5)}
                                    
  beamEnum2Text=function(enum) {enum=as.character(enum)
    return(switch(enum,
                                      '1' = 'Biosonics',
                                      '3' = 'Furuno',
                                      '2' = 'Hti',
                                      '4' = 'None',
                                      '0' = 'SimradLobe',
                                      '5' = 'Sonic'))}
  #find virtual varialble
  finder=EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=finder$msg #we will keep track of messages produced by the function here
  acoVar=finder$EVVar #this is the COM address of the acoustic virtual variable of interest
  #check acoustic variable type
  varType=acoVar$Operator()
  if(!(varType %in% c(75,114))){
     msg=paste(Sys.time(),' : input variable ',acoVarName,' is not a split-beam type (method 1 and 2) virtual acoustic variable')
     warning(msg)
     msgV=c(msgV,msg)
     return(list(settings=NULL,msg=msgV))
  }
  #get current settings
  
  STparametersDataFrame=data.frame(matrix(NA,2,4))
  nameV=c('BeamCompensationModel','MaximumBeamCompensation','MaximumStdDevOfMajorAxisAngles',
          'MaximumStdDevOfMinorAxisAngles')
  names(STparametersDataFrame)=nameV
  
  row.names(STparametersDataFrame)=c('pre','post')
  #Single target detection parameters settings:
  singleTargetPars=acoVar[['Properties']][['SingleTargetDetectionSplitBeamParameters']]
  print(singleTargetPars$BeamCompensationModel())
  STparametersDataFrame$BeamCompensationModel[1]=beamEnum2Text(singleTargetPars$BeamCompensationModel())
  STparametersDataFrame$MaximumBeamCompensation[1]=singleTargetPars$MaximumBeamCompensation()
  
  STparametersDataFrame$MaximumStdDevOfMajorAxisAngles[1]=singleTargetPars$MaximumStdDevOfMajorAxisAngles()
  STparametersDataFrame$MaximumStdDevOfMinorAxisAngles[1]=singleTargetPars$MaximumStdDevOfMinorAxisAngles()
  
  
  if(verbose){
    message('---------- PRE-CHANGE SPLIT-BEAM SINGLE TARGET PARAMETERS ---------------')  
    message('Beam compensation model = ', STparametersDataFrame$BeamCompensationModel[1])
    message('Maximum beam compensation = ', STparametersDataFrame$MaximumBeamCompensation[1],' dB')
    message('Maximum SD of Major Axis Angles = ', STparametersDataFrame$MaximumStdDevOfMajorAxisAngles[1],'deg')
    message('Maximum pulse length = ', STparametersDataFrame$MaximumStdDevOfMinorAxisAngles[1],'deg')
    message('------------------------------------------------------------------------')  }
  
  #Change settings:
  msg=paste(Sys.time(),' : Changing SPLIT-BEAM single target detection settings in viritual variable ',acoVarName)
  message(msg)
  msgV=c(msgV,msg)
  
  if(!is.null(BeamCompensationModel)) singleTargetPars[['BeamCompensationModel']]=beamText2enum(BeamCompensationModel)
  if(!is.null(MaximumBeamCompensation)) singleTargetPars[['MaximumBeamCompensation']]=MaximumBeamCompensation
  if(!is.null(MaximumStdDevOfMajorAxisAngles)) singleTargetPars[['MaximumStdDevOfMajorAxisAngles']]=MaximumStdDevOfMajorAxisAngles
  if(!is.null(MaximumStdDevOfMinorAxisAngles)) singleTargetPars[['MaximumStdDevOfMinorAxisAngles']]=MaximumStdDevOfMinorAxisAngles
  
  #change settings were changed:
  singleTargetPars=acoVar[['Properties']][['SingleTargetDetectionSplitBeamParameters']]
  STparametersDataFrame$BeamCompensationModel[2]=beamEnum2Text(singleTargetPars$BeamCompensationModel())
  STparametersDataFrame$MaximumBeamCompensation[2]=singleTargetPars$MaximumBeamCompensation()
  
  STparametersDataFrame$MaximumStdDevOfMajorAxisAngles[2]=singleTargetPars$MaximumStdDevOfMajorAxisAngles()
  STparametersDataFrame$MaximumStdDevOfMinorAxisAngles[2]=singleTargetPars$MaximumStdDevOfMinorAxisAngles()
  
  if(verbose){
    message('----------- POST-CHANGE SPLIT-BEAM SINGLE TARGET PARAMETERS ---------------')  
    message('Beam compensation model = ', STparametersDataFrame$BeamCompensationModel[2])
    message('Maximum beam compensation = ', STparametersDataFrame$MaximumBeamCompensation[2],' dB')
    message('Maximum SD of Major Axis Angles = ', STparametersDataFrame$MaximumStdDevOfMajorAxisAngles[2],'deg')
    message('Maximum SD of Minor Axis Angles = ', STparametersDataFrame$MaximumStdDevOfMinorAxisAngles[2],'deg')
    
    message('------------------------------------------------------------------------')  }
  
  STparametersDataFrame=data.frame(t(STparametersDataFrame))
 
  row.names(STparametersDataFrame)=nameV
  names(STparametersDataFrame)=c('pre','post')

  return(list(settings=STparametersDataFrame,msg=msgV))
}


