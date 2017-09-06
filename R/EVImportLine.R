#' Import an Echoview line (EVL format) file
#' 
#'This function imports an Echoview line (.EVL) format line into Echoview and can rename the imported line or replace a line.
#'  
#'@import RDCOMClient
#'@param EVFile ("COMIDispatch) An Echoview file COM object 
#'@param pathAndFn = NULL (character) path and file name of an EVL file
#'@param lineName = NULL (character) desired line name
#'@param existingLine = NULL behaviour if a line with the same name as lineName already exists in Echoview.  See details.
#'@export
#'@return list $lineObj=COM object for the imported line (class COMIDispatch), msg= vector of messages)
EVImportLine <- function (EVFile, pathAndFn=NULL,lineName=NULL,existingLine=NULL) {
  
  if(is.null(pathAndFn))
  {msg=paste(Sys.time(),' : No path and filename specified in EVImportLine(pathAndFn)',sep='')
  warning(msg) 
  return(list(msg=msg))
  }
  nbr.of.lines.pre.import<-EVFile[['Lines']]$Count()  
  msgV=paste(Sys.time(),' : Importing line',pathAndFn,sep='')
  message(msgV)
  importedLineBoolean<-EVFile$Import(pathAndFn)
  if(importedLineBoolean)
  {
    msg=paste(Sys.time(),' : Successfully imported ',pathAndFn,sep='')
    message(msg)
    msgV=c(msgV,msg)
  } else{
    msg=paste(Sys.time(),' : Failed to import ',pathAndFn,sep='')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(msg=msgV))
  }  
  #find newline object
  importedLineObj=EVFile[['Lines']]$Item(nbr.of.lines.pre.import)
  
  if(is.null(lineName)) {
    msg=paste(Sys.time(),' : Imported line object name = ',importedLineObj$Name(),sep='')
    message(msg)
    msgV=c(msgV,msg)
    invisible(list(lineObj=importedLineObj,msg=msgV))
  }else{
    msg=paste(Sys.time(),' : Renaming imported line object name from ',importedLineObj$Name(),
              ' to ',lineName,sep='')
    message(msg)
    msgV=c(msgV,msg)
    #check if there is already a line with the new name
    line.check <- EVFile[["Lines"]]$FindByName(lineName)
    if (is.null(line.check) == FALSE) {
      if(is.null(existingLine)){
        msg=paste(Sys.time(), "Error: A line already exists with the name", lineName, 
                  ". Line not renamed. New Line is called ",importedLineObj$Name(), sep = " ")
        warning(msg)
        return(list(lineObj=importedLineObj,msg=msgV))
      }else{
        if(existingLine==0){
          EVDeleteLine(EVFile, EVFindLineByName(EVFile, 'aaa'))
          msg = paste(Sys.time(), "Line existed prior to import and has been deleted")
        }else{
          EVRenameLine(EVFile, line.check,existingLine)
          msg = paste(Sys.time(), "Line existed prior to import. The old line has been renamed to",existingLine)
        }
        warning(msg)
        msgV=c(msgV,msg)
      }
    }
    importedLineObj[["Name"]] <- lineName
    #check that line has been renamed
    
    if (importedLineObj$Name() == lineName) {
      msg=paste(Sys.time(), ": Success, line renamed as", lineName, sep = " ")
      message(msg)
      msgV=c(msgV,msg)
      invisible(list(lineObj=importedLineObj,msg=msgV))
    } else {
      msg=paste(Sys.time(), ": Could not rename line")
      warning(msg)
      msgV=c(msgV,msg)
      return(list(lineObj=importedLineObj,msg=msgV))
    }  
    #end of rename line if statement.
  }
}
