#' Create an Echoview EVR (Region definition) file from depth and time intervals 
#' 
#' This function generates Echoview Region definitions according to predefined starting and end depts at given times.
#' The results are save as .EVR file for import into Echoview
#' 
#' @param depths list of depths (can be multiple regions, where each is a list element) 
#' @param times (character) list of times in HHMMSS
#' @param dates  list of dates as character \%Y\%m\%d (can be multiple regions, where each is a list element)
#' @param rName  = "Region" (character) Name of the region
#' @param rClass = "Selection".  character Define class, otherwise it will be Unclassified
#' @param rType = (integer; 0 to 4) see details
#' @param dir = NULL character Path to the output folder
#' @param fn = NULL character Name of the output filename
#' @param rNotes  = list("") character Region Notes
#' @return saves an Echoview EVR (Region definition) files
#' @keywords Echoview com scripting
#' @details The following region types (integer 0 to 4) are available in Echoview 0 - Bad Data (no data); 1 - Analysis; 2 - Marker; 3 - fishtracks; 4 - Bad Data (empty water)
#' @export
#' @examples
#' \dontrun{
#' # Minimal example generating 2 regions
#' depths <- list(c(10,35,50,74,82), c(15,250,300,350,60))
#' times <-  list(c("090000","091000","091500","091000","090300"),
#'                c("093000","094000","094500","094000","093300"))
#' dates <- list(c(rep("20170901",5)),c(rep("20170901",5)))
#' dir <- 'F:\\'
#' fn <- 'imaginaryRegions'
#' makeEVRfromDT(depths,times,dates,fn=fn,dir=dir)
#'  }

makeEVRfromDT<-function(depths,times, dates, rName = "Region", rClass = "Selection",rType = 1,
                        dir=NULL, 
                        fn=NULL, 
                        rNotes=list("")) 
{   
  if(!all(c(length(depthStart),length(depthStop),length(timeStart))==length(timeStop)))
    stop('ARGS depths and times must have the same length.')
  opFile=paste(fn,'.evr',sep='')
  if(!is.null(dir))
    opFile=paste(dir,opFile,sep='')
  
  #number of regions
  nreg <- length(dates)
  
  ###HEADER
  cat(paste('EVRG 7 8.0.91.31697'),file=opFile,append=FALSE, sep='\n')
  cat(paste(nreg,'\n'),file=opFile, append=TRUE, sep ='\n')
  ##########
  
  #####FOR EACH REGION
  for(i in 1:nreg){
    message(paste(Sys.time(),"Processing region", i))
    if(length(rNotes)>=i){if(nchar(rNotes[[i]])>0){linesNotes = 1}else{linesNotes=0}}
    linesDetec = 0
    
    #min point
    min_ind <- which(as.numeric(times[[i]]) == min(as.numeric(times[[i]])))
    #max point
    max_ind <- which(as.numeric(times[[i]]) == max(as.numeric(times[[i]])))
    
    
    #####FIRST LINE
    cat(paste(13, #Region Struxture Version (currently 13)
              length(dates[[i]]), #Number of points defining the shape
              i, #region id
              0,
              2,
              -1, #Dummy -1
              1, #valid poits 1, otherwise 0
              dates[[i]][min_ind],
              substr(paste0(as.character(times[[i]][min_ind]),'0000'),1,10),
              formatC(depths[[i]][min_ind],digits=6,format='f'),
              dates[[i]][max_ind],
              substr(paste0(as.character(times[[i]][max_ind]),'0000'),1,10),
              depths[[i]][max_ind]),file=opFile,append=TRUE,sep='\n')
    
    cat(paste(linesNotes),file=opFile,append=TRUE,sep='\n')
    cat(paste(linesDetec),file=opFile,append=TRUE,sep='\n')
    cat(paste(rClass),file=opFile,append=TRUE,sep='\n')
    
    pts = ""
    for(d in 1:length(dates[[i]])){
      pts <- append(pts,paste(dates[[i]][d],
                              substr(paste0(as.character(times[[i]][d]),'0000'),1,10),
                              formatC(depths[[i]][d],digits=6,format='f')))
    }
    cat(c(pts[2:length(pts)],i,'\n'), file=opFile,append=TRUE)
    cat(paste0(rName,i,'\n'),file=opFile,append=TRUE,sep="\n")
  }
  
  message(paste(Sys.time(), 'Echoview region (.evr) file written to: ',file=opFile))
}  
