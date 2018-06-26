#' Broadband Single target detection
#' 
#' This function creates a Wideband Single Target Detection Variable
#' 
#' @param EVFile An Echoview file COM object
#' @param EVVar An Echoview Variable, accepts inputs as Character, list or Variable object (COMIDispatch)
#' @param Operand2 An Echoview Variable, accepts inputs as Character, list or Variable object (COMIDispatch)
#' @param TsThreshold numeric [dB] Single target detection threshold
#' @param MinimumPulseLength numeric Minimum normalized pulse length 
#' @param MaximumPulseLength numeric Maximum normalized pulse length ,
#' @param BeamCompensationModel numeric [1,4] 1 = Simrad Lobe, 4 = None,
#' @param MaximumBeamCompensation  NULL=,numeric Maximum beam compensationngles = NULL,
#' @param MaximumStdDevOfMinorAxisAngles Maximum standard deviation of minor axis angles
#' @param MaximumStdDevOfMajorAxisAngles Maximum standard deviation of major axis angles
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' #To be added - Needs Example data
#' #Starting Echoview
#' echoview = StartEchoview()
#' #Create a new EV File
#' EVFile <- EVCreateNew(EVAppObj=echoview, dataFiles = "rawfile.raw")$EVFile
#' #Set the First and second Operand
#' EVVar <- "Fileset 1: TS pulse compressed wideband pings T1"
#' Operand2 <- "Fileset 1: angular position pulse compressed wideband pings T1"
#' #Create a wideband single target detection varable
#' ST <- EVSTWideband(EVFile=EVFile,EVVar=EVVar,Operand2=Operand2)
#' # Change TS detection threshold
#' ST <- EVSTWideband(EVFile=EVFile,EVVar=EVVar,Operand2=Operand2,TsThreshold=-90)
#' 


EVSTWideband <- function(EVFile, EVVar, Operand2, 
                       TsThreshold = NULL,
                       MinimumPulseLength = NULL,
                       MaximumPulseLength = NULL,
                       BeamCompensationModel = NULL,
                       MaximumBeamCompensation = NULL,
                       MaximumStdDevOfMajorAxisAngles = NULL,
                       MaximumStdDevOfMinorAxisAngles = NULL){
  
  #Test class of EVVar to see how it should be used
  EVVar <- switch(class(EVVar),
                "character" = {EVAcoVarNameFinder(EVFile, acoVarName=EVVar)$EVVar},
                "list" = EVVar$EVVar,
                "COMIDispatch" = EVVar)
  #Test class of Operand2 to see how it should be used
  Operand2 <- switch(class(Operand2),
                     "character" = {EVAcoVarNameFinder(EVFile, acoVarName=Operand2)$EVVar},
                     "list" = Operand2$EVVar,
                     "COMIDispatch" = Operand2)
  #Create Wideband SIngle Target detaection variable using enum 143
  message(paste0(Sys.time(),": Creating Wideband Single Target Detection Object"))
  ST <- EVVar$AddVariable(143)
  ST$SetOperand(2,Operand2)
  
  #Get the parent objects of the three settings groups
  STprop <- ST[["Properties"]][["SingleTargetDetectionParameters"]]
  STWBprop <- ST[["Properties"]][["SingleTargetDetectionWidebandParameters"]]
  STFilter <- ST[["Properties"]][["FilterTargets"]]
  
  #Set custom settings for single target detection
  if(!is.null(TsThreshold)){STprop[["TsThreshold"]] = TsThreshold}
  if(!is.null(MinimumPulseLength)){STprop[["MinimumPulseLength"]] = MinimumPulseLength}
  if(!is.null(MaximumPulseLength)){STprop[["MaximumPulseLength"]] = MaximumPulseLengt}
  if(!is.null(BeamCompensationModel)){STWBprop[["BeamCompensationModel"]] = BeamCompensationModel}
  if(!is.null(MaximumBeamCompensation)){STWBprop[["MaximumBeamCompensation"]] = MaximumBeamCompensation}
  if(!is.null(MaximumStdDevOfMajorAxisAngles)){STWBprop[["MaximumStdDevOfMajorAxisAngles"]] = MaximumStdDevOfMajorAxisAngles}
  if(!is.null(MaximumStdDevOfMinorAxisAngles)){STWBprop[["MaximumStdDevOfMinorAxisAngles"]] = MaximumStdDevOfMinorAxisAngles}
  
  message(paste0(Sys.time(),": Success: Created Wideband Single Target Detection Object"))
}

