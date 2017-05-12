#' Import an Echoview region definitions file (.evr)
#' 
#' This function imports a region definitions file (.evr) using COM scripting
#' @param EVFile An Echoview file COM object
#' @param evrFile An Echoview region definitions file (.evr) path and name
#' @return If successful, A list object with three elements. $regionCOMObjs a list comprising of the COM object of each imported region;
#' $regionMeta metadata of the regions imported: the region ID, region name and region class, and $msg message for processing log
#' @keywords Echoview COM scripting 
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj,'~~/KAOS/KAOStemplate.EV')$EVFile
#' EVImportRegionDef(EVFile = EVFile, evrFile = '~~/KAOS/off transect regions/20030114_1200000000.evr')
#'}

EVImportRegionDef <- function (EVFile, evrFile) {
  
  if(!file.exists(evrFile)) {msg=paste(Sys.time(),': EVR file not found.')
                            warning(msg)
                            return(list(regionCOMObjs=NA,regionMeta=NA,msg=msg))}
  pren=EVFile[['Regions']]$Count()#nbr of regions pre load
  importOKFlag=EVFile$Import(evrFile)
  if(!importOKFlag){ #check import
    msg=paste(Sys.time(),': An unknown error occured during the import of',evrFile)
    warning(msg)
    invisible(list(regionCOMObjs=NA,regionMeta=NA,msg=msg))
  }
  postn=EVFile[['Regions']]$Count()#nbr of regions pre load
  addedn=postn-pren
  if(addedn==0) {msg=paste(Sys.time(),' : No regions were added.  Check evr file',evrFile)
                warning(msg)
                return(list(regionCOMObjs=NA,regionMeta=NA,msg=msg))} else {
                msgV=  paste(Sys.time(),':',addedn,'Regions were added from file',evrFile)
                message(msgV)
                }
  regionIND=((pren+1):postn)-1 #index of regions for use with the Count() EV COM method
  regionCOML=vector(length=addedn,mode='list')
  rIDV=rNameV=rClassV=vector(length=addedn)
  for(i in 1:addedn)
  {
    regionCOML[[i]]=EVFile[['Regions']]$Item(regionIND[i])
    rIDV[i]=regionCOML[[i]]$ID()
    rNameV[i]=regionCOML[[i]]$Name()
    rClassV[i]=regionCOML[[i]]$RegionClass()$Name()
  }
  invisible(list(regionCOMObjs=regionCOML,regionMeta=data.frame(id=rIDV,name=rNameV,class=rClassV),
                 msg=msgV))
  }  
