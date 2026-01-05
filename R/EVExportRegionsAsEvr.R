#' Export Echoview region definitions to (an .evr) file
#'
#' Exports region definitions from an open Echoview file to an Echoview
#' region definition file (\code{.evr}). Regions can be exported either
#' for a single region classification or for all region classifications
#' contained in the Echoview file.
#'
#' This function uses the Echoview COM interface and therefore requires
#' Echoview to be installed and running on a Windows system.
#'
#' @param EVFile COM object.
#'   An active Echoview file COM object, typically created using
#'   \code{EVOpenFile()} or similar EchoviewR functions.
#'
#' @param className character.
#'   The name of the Echoview region classification to export.
#'   Use \code{"all"} to export regions from all region classifications.
#'
#' @param filePath character.
#'   Full file path for the output Echoview region definition file
#'   (usually with extension \code{.evr}).
#'
#' @return Invisibly returns a list with element \code{msg}, containing
#'   a character vector of timestamped status messages describing the
#'   export process.
#'
#' @details
#' When \code{className = "all"}, the Echoview COM method
#' \code{ExportDefinitionsAll()} is called. Otherwise,
#' \code{ExportDefinitions()} is used to export only the specified
#' region classification.
#'
#' If the export fails, a warning is issued.
#'
#' @note This function was tested with Echoview version 15.1.75.
#'#' @examples
#' \dontrun{
#' # Export all region classifications
#' EVExportRegionsAsEvr(
#'   EVFile   = evfile,
#'   className = "all",
#'   filePath = "C:/acoustics/regions_all.evr"
#' )
#'
#' # Export a single region classification
#' EVExportRegionsAsEvr(
#'   EVFile   = evfile,
#'   className = "Krill regions",
#'   filePath = "C:/acoustics/krill_regions.evr"
#' )
#' }
#'
#' @export
EVExportRegionsAsEvr=function(EVFile,className,filePath)
{
  msgV=paste(Sys.time(),': Attempting to export region class called',className,'to file:',filePath)
  message(msgV)
  region_obj=EVFile[['Regions']]
  if(className=='all'){
    msg=paste(Sys.time(),' : exporting all regions (from all region calssifications.)')
    message(msg)
    msgV=c(msgV,msg)
    didit=region_obj$ExportDefinitionsAll(filePath)} else {
      msg=paste(Sys.time(),' : exporting regions for region classification:',className)
      message(msg)
      msgV=c(msgV,msg)
      didit=region_obj$ExportDefinitions(filePath,className)
    }
  if(didit){
    msg=paste(Sys.time(),' : exported regions to',filePath)
    message(msg)
    msgV=c(msgV,msg)
  } else {
    msg=paste(Sys.time(),' : Failed to export regions to',filePath)
    warning(msg)
    msgV=c(msgV,msg)
  }
  invisible(list(msg=msgV))
}
