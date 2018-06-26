#' Export Wideband Single Targets Frequency Response
#' 
#' This function exports the broadband single targets frequency response
#' 
#' @param EVFile An Echoview file COM object
#' @param STVar An Wideband Single Target Detection Echoview Variable, accepts inputs as Character, list or Variable object (COMIDispatch)
#' @param outfn FIlename for the TS export
#' @param MinMax Include Min and Max
#' @param WindowSize Set the WIndow Size
#' @param WindowUnit Unit of the window 0 = Meters, 1 = Pulse Lengths, 2 = Samples
#' @param EVVar Specifies a TS pulse compressed wideband variable for frequency response calculations.
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
#' #Export the single target frequency response
#' EVExportWBST_FR(ST,"WB_ST_FR.csv",0,0,1,2,EVVar)
#' 

EVExportWBST_FR <- function(EVFile=EVFile,
                            STVar, 
                            outfn,
                            AverageResults=0,
                            MinMax = 0,
                            WindowSize = 1,
                            WindowUnit = 2,
                            EVVar){
  
  #Test class of EVVar to see how it should be used
  StVar <- switch(class(StVar),
                  "character" = {EVAcoVarNameFinder(EVFile, acoVarName=StVar)$EVVar},
                  "list" = StVar$EVVar,
                  "COMIDispatch" = StVar)
  
  EVVar <- switch(class(EVVar),
                  "character" = {EVAcoVarNameFinder(EVFile, acoVarName=EVVar)$EVVar},
                  "list" = EVVar$EVVar,
                  "COMIDispatch" = EVVar)
  message(paste0(Sys.time(),": Exporting Single Targets"))
  export <- StVar$ExportSingleTargetWidebandFrequencyResponse(
    outfn, 
    0,0,.4,0,EVVar)
  if(export==TRUE){message(paste0(Sys.time(),": Export completed."))}
  if(export==FALSE){message(paste0(Sys.time(),": Export failed. Possibly the wrong variables are used as an input."))}
  
}