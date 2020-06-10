#' Bottom Detection
#' 
#' This function runs a bottom detection on an echogram
#' Bottom algorithms and settings are explained in the [Echoview Help File]('https://support.echoview.com/WebHelp/Reference/Algorithms/Line_picking_algorithm.htm')
#' 
#' @param EVFile An Echoview file COM object
#' @param EVVar An Echoview Variable, accepts inputs as Character, list or Variable object (COMIDispatch)
#' @param LineName Character of the output name for the detected Line
#' @param algorithm numeric [0 - 2] Defines which bottom detection algorithm should be used: 0 for Delta Sv, 1 for Maximum Sv, 2 for Best bottom Candidate
#' @param StartDepth numeric [m] Minimum bottom detection depth
#' @param StopDepth numeric [m] maximum bottom detection detpth
#' @param MinSv numeric [dB] minimum detection Sv
#' @param UseBackstep Boolean [True or False]
#' @param BackstepRange numeric [m] Backstep range
#' @param DiscriminationLevel numeric [dB] Minimum discrimination threshold
#' @param MaxDropouts numeric [samples] Maximum number of dropout samples before bottom detection fails
#' @param WindowRadius numeric [samples] Search window size
#' @param PeakThreshold numeric [dB] Threshold for peak detection
#' @param MinPeakAssymmetry numeric
#' @param replaceOldBottom Boolean (TRUE or FALSE) If TRUE and a line with the same name as LineName already exists, the old line will be overwritten with the new one
#' @param SpanGaps Boolean (TRUE or FALSE) Decides wheteher span gaps is activated or notm default=TRUE
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#' \dontrun{
#' #To be added - Needs Example data
#' #Starting Echoview
#' echoview = StartEchoview()
#' #Create a new EV File
#' EVFile <- EVCreateNew(EVAppObj=echoview, dataFiles = "rawfile.raw")$EVFile
#' Varname <- "Fileset 1: Sv pulse compressed wideband pings T1"
#' bottom <- EVBottomDetection(EVFile, EVVar=Varname, LineName="Bottom")
#' # Change the algorithm to Best bottom candidate
#' bottom <- EVBottomDetection(EVFile, EVVar=Varname, LineName="Bottom",algorithm=2)
#' }
#' 

EVBottomDetection <- function(EVFile, EVVar, LineName="Bottom",
                              algorithm = NULL, #2,
                              StartDepth = NULL, #5,
                              StopDepth = NULL, #2000,
                              MinSv = NULL, #-50,
                              UseBackstep = NULL, #"True",
                              DiscriminationLevel = NULL,
                              BackstepRange = NULL, #-0.50,
                              PeakThreshold = NULL,
                              MaxDropouts = NULL, #2,
                              WindowRadius = NULL, #8,
                              MinPeakAssymmetry = NULL, #-1.0,
                              replaceOldBottom = TRUE,
                              SpanGaps = TRUE){

  #Test class of EVVar to see how it should be used
  EVVar <- switch(class(EVVar),
       "character" = {EVAcoVarNameFinder(EVFile, acoVarName=EVVar)$EVVar},
       "list" = EVVar$EVVar,
       "COMIDispatch" = EVVar)
  
  #Set Bottom detection settings
  setlp <- EVFile[["Properties"]][["LinePick"]]
  
  if(!is.null(algorithm)){setlp[["Algorithm"]] <- algorithm} # 0 for Delta Sv, 1 for Maximum Sv, 2 for Best bottom Candidate
  if(!is.null(StartDepth)){setlp[["StartDepth"]] <- StartDepth}
  if(!is.null(StopDepth)){setlp[["StopDepth"]] <- StopDepth}
  if(!is.null(MinSv)){setlp[["MinSv"]] <- MinSv}
  if(!is.null(UseBackstep)){setlp[["UseBackstep"]] <- UseBackstep}
  if(!is.null(DiscriminationLevel)){setlp[["DiscriminationLevel"]] <- DiscriminationLevel}
  if(!is.null(BackstepRange)){setlp[["BackstepRange"]] <- BackstepRange}
  if(!is.null(PeakThreshold)){setlp[["PeakThreshold"]] <- PeakThreshold}
  if(!is.null(MaxDropouts)){setlp[["MaxDropouts"]] <- MaxDropouts}
  if(!is.null(WindowRadius)){setlp[["WindowRadius"]] <- WindowRadius}
  if(!is.null(MinPeakAssymmetry)){setlp[["MinPeakAsymmetry"]] <- MinPeakAssymmetry}
  
  #Detect Bottom
  message(paste0(Sys.time(),": Detecting Bottom..."))
  #Detect bottom with or without span gaps
  if(SpanGaps){
    bottom<- EVFile[["Lines"]]$CreateLinePick(EVVar,T)
  }else{
      bottom<- EVFile[["Lines"]]$CreateLinePick(EVVar,F)
      }
  
  message(paste0(Sys.time(),": Success Bottom detected successfully..."))  
  #Overwrite old bottom line (if needed)
  if(replaceOldBottom){
    message(paste0(Sys.time(), ": Checking if Line with name ", LineName, " already exists"))
    oldbottom <- EVFindLineByName(EVFile,LineName)
    if(!is.null(oldbottom)){
      message(paste0(Sys.time(),": Bottom of name ", LineName, " detected."))
      bottom <- oldbottom$OverwriteWith(bottom)
      message(paste0(Sys.time(), ": Old bottom has been overwritten..."))
    }
  }
  
  bottom[['Name']] <- LineName
  message(paste0(Sys.time(),": Success new bottom line ", LineName, " has been created"))
}

  
