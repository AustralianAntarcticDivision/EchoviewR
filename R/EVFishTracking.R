#' Detect Fish Tracks
#' 
#' This function runs fish track detection on an single target echogram
#' Fish track algorithms and settings are explained in the [Echoview Help File]('https://support.echoview.com/WebHelp/Reference/Algorithms/Line_picking_algorithm.htm')
#'
#' @param EVFile An Echoview file COM object
#' @param EVVar An Echoview single target variable as a character string
#' @param FishTrackRegionClass string indicating the region class name for the fish tracks to be assigned
#' @param export full file path as a text string for an export file (i.e. "C:/FishTracking/Exports/FishTracks.csv"). NULL value skips export of fish tracks. 
#' @param DataDimensions numeric Determine which algorithim is used for fish tracks. Accepted values are 2 (single or dual beam) or 4 (split beam)
#' @param Alpha.MajorAxis numeric Increases Alpha gain setting of Major Axis
#' @param Alpha.MinorAxis numeric Increases Alpha gain setting of Minor Axis
#' @param Alpha.Range numeric Increases Alpha gain setting of Range
#' @param Beta.MajorAxis numeric Increases Beta gain setting of Major Axis
#' @param Beta.MinorAxis numeric Increases Beta gain setting of Minor Axis
#' @param Beta.Range numeric Increases Beta gain setting of Range
#' @param ExclusionDistance.MajorAxis numeric Increases distance value of major axis in target gate parameters
#' @param ExclusionDistance.MinorAxis numeric Increases distance value of minor axis in target gate parameters
#' @param ExclusionDistance.Range numeric Increases distance value of range in target gate parameters
#' @param MissedPingExpansion.MajorAxis numeric Increases expansion value of major axis in target gate parameters
#' @param MissedPingExpansion.MinorAxis numeric Increases expansion value ofminor axis in target gate parameters 
#' @param MissedPingExpansion.Range numeric Increases expansion value of missed pings range in target gate parameters. 0 disables missed ping expansion
#' @param Weights.MajorAxis numeric 1-100 Assigns the weighting of the major axis
#' @param Weights.MinorAxis numeric 1-100 Assigns the weighting of the minor axis
#' @param Weights.Range numeric 1-100 Assigns the weighting of the range
#' @param Weights.TS numeric 1-100 Assigns the weighting of the target strength
#' @param Weights.PingGap numeric 1-100 
#' @param MinimumTargets numeric Minimum number of single targets to be considered a fish track
#' @param MinimumPings numeric Maximum number of pings in a track
#' @param MaximumGap numeric Maximum ping gap between tracks
#'
#' @return
#' @export
#'
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVFishTracking(EVFile = EvFile,
#' EVVar = "Single target detection - wideband 1",
#' FishTrackRegionClass = "FishTracks",
#' MinimumTargets = 5,
#' export = "C:/FishTracking/Exports/FishTracks.csv")
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
  
  # Detect fish tracks
  # DetectFishTracks returns number of tracks if successful
  
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
  
  # if export file has been provided, export the fish tracks
  if(!is.null(export)){
      FileExported <- EVVar$ExportFishTracksByRegions(
      export, 
      EVFile[["RegionClasses"]]$FindByName(FishTrackRegionClass))
      
      # FileExported returns TRUE if export was successful
      if(FileExported){
      msg <- paste(Sys.time(), ' : ', 'Fish tracks from ', FishTrackRegionClass, 
                   " was successfully exported as ", export, sep = '')
      message(msg)
      }
      # Provide alternative error message if FALSE
      if(!FileExported){
        warning("Fish tracks were not exported.")
      }
  }
}
