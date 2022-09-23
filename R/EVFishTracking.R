#' Detect Fish Tracks
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
#' @param MaxDropOuts numeric [samples] Maximum number of dropout samples before bottom detection fails
#' @param windowRadius numeric [samples] Search window size
#' @param PeakThreshold numeric [dB] Threshold for peak detection
#' @param MinPeakAssymmetry numeric
#' @param replaceOldBottom Boolean (TRUE or FALSE) If TRUE and a line with the same name as LineName already exists, the old line will be overwritten with the new one
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

EVFishTracking <- function(EVFile, 
                           EVVar, 
                           FishTrackRegionClass = "Unclassified regions",
                           export = NULL,
                           DataDimensions = NULL,#4
                           Alpha.MajorAxis = NULL,#0.1
                           Alpha.MinorAxis = NULL,#0.2
                           Alpha.Range = NULL,#0.3
                           Beta.MajorAxis = NULL,#0.4
                           Beta.MinorAxis = NULL,#0.5
                           Beta.Range = NULL,#0.6
                           ExclusionDistance.MajorAxis = NULL,#0.7
                           ExclusionDistance.MinorAxis = NULL,#0.8
                           ExclusionDistance.Range = NULL,#0.9
                           MissedPingExpansion.MajorAxis = NULL,#1.0
                           MissedPingExpansion.MinorAxis = NULL,#1.1
                           MissedPingExpansion.Range = NULL,#1.2
                           Weights.MajorAxis = NULL,#1.3
                           Weights.MinorAxis = NULL,#1.4
                           Weights.Range = NULL,#1.5
                           Weights.TS = NULL,#1.6
                           Weights.PingGap = NULL,#1.7
                           MinimumTargets = NULL,#4
                           MinimumPings = NULL,#5
                           MaximumGap = NULL#3
                           ){
  
  #Test class of EVVar to see how it should be used
  EVVar <- switch(class(EVVar),
                  "character" = {EVAcoVarNameFinder(EVFile, acoVarName=EVVar)$EVVar},
                  "list" = EVVar$EVVar,
                  "COMIDispatch" = EVVar)
  
  #Set Fish tracking settings
  setlp <- EVFile[["Properties"]][["FishTracking"]]
  if(!is.null(DataDimensions)){setlp[["DataDimensions"]] <- DataDimensions} 
  
  alpha<-setlp[["Alpha"]]
  if(!is.null(Alpha.MajorAxis)){alpha[["MajorAxis"]] <- Alpha.MajorAxis} 
  if(!is.null(Alpha.MinorAxis)){alpha[["MinorAxis"]] <- Alpha.MinorAxis}
  if(!is.null(Alpha.Range)){alpha[["Range"]] <- Alpha.Range}
  
  beta<-setlp[["Beta"]]
  if(!is.null(Beta.MajorAxis)){beta[["MajorAxis"]] <- Beta.MajorAxis}
  if(!is.null(Beta.MinorAxis)){beta[["MinorAxis"]] <- Beta.MinorAxis}
  if(!is.null(Beta.Range)){beta[["Range"]] <- Beta.Range}
  
  exclusion <- setlp[["ExclusionDistance"]]
  if(!is.null(ExclusionDistance.MajorAxis)){exclusion[["MajorAxis"]] <- ExclusionDistance.MajorAxis}
  if(!is.null(ExclusionDistance.MinorAxis)){exclusion[["MinorAxis"]] <- ExclusionDistance.MinorAxis}
  if(!is.null(ExclusionDistance.Range)){exclusion[["Range"]] <- ExclusionDistance.Range}
  
  missed = setlp[["MissedPingExpansion"]]
  if(!is.null(MissedPingExpansion.MajorAxis)){missed[["MajorAxis"]] <- MissedPingExpansion.MajorAxis}
  if(!is.null(MissedPingExpansion.Range)){missed[["Range"]] <-MissedPingExpansion.Range}
  
  weights <- setlp[["Weights"]]
  if(!is.null(Weights.MajorAxis)){weights[["MajorAxis"]] <- Weights.MajorAxis}
  if(!is.null(Weights.MinorAxis)){weights[["MinorAxis"]] <- Weights.MinorAxis}
  if(!is.null(Weights.Range)){weights[["Range"]] <- Weights.Range}
  if(!is.null(Weights.TS)){weights[["TS"]] <- Weights.TS}
  if(!is.null(Weights.PingGap)){weights[["PingGap"]] <- Weights.PingGap}  
  if(!is.null(MinimumTargets)){setlp[["MinimumTargets"]] <- MinimumTargets}
  if(!is.null(MinimumPings)){setlp[["MinimumPings"]] <- MinimumPings}
  if(!is.null(MaximumGap)){setlp[["MaximumGap"]] <- MaximumGap}
  

  Ntracks <- EVVar$DetectFishTracks(FishTrackRegionClass)
  if(Ntracks == 0) {
    msg <- paste(Sys.time(), ' : No fish tracks detected.', sep = '')
    warning(msg)
  } else {
    msg <- paste(Sys.time(), ' : ', Ntracks,
                 " fish tracks were detected and assigned to ",
                 FishTrackRegionClass, " Region Class.", sep = '')
    message(msg)
  }
  
  if(!is.null(export)){
    # This is a temporary fix - I can't get $ExportFishTracksByRegions to work.
    FileExported <- EVVar$ExportFishTracksByRegionsAll(export)
    if(FileExported){
      msg <- paste(Sys.time(), ' : ', export, " was successfully exported.", sep = '')
      message(msg)
    }
  }
}