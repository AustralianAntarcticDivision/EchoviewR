#' Import an Echoview line (EVL format) file
#' 
#'This function imports an Echoview line (.EVL) format line into Echoview and can rename the imported line or replace a line.
#'  
#'@import RDCOMClient
#'@param EVFile ("COMIDispatch) An Echoview file COM object 
#'@param pathAndFn = NULL (character) path and file name of an EVL file
#'@param lineName = NULL (character) desired line name
#'@param overwriteExistingLine = TRUE (boolean) overwrite an existing line named lineName
#'@note Overwriting an existing line only works with editable lines.
#'@export
#'@return list $lineObj=COM object for the imported line (class COMIDispatch), msg= vector of messages)
EVImportLine=function (EVFile, pathAndFn = NULL, lineName = NULL, overwriteExistingLine = TRUE) 
  {
    if (is.null(pathAndFn)) {
      msg = paste(Sys.time(), " : No path and filename specified in EVImportLine(pathAndFn)", 
                  sep = "")
      warning(msg)
      return(list(msg = msg))
    }
    nbr.of.lines.pre.import <- EVFile[["Lines"]]$Count()
    msgV = paste(Sys.time(), " : Importing line", pathAndFn, 
                 sep = "")
    message(msgV)
    importedLineBoolean <- EVFile$Import(pathAndFn)
    if (importedLineBoolean) {
      msg = paste(Sys.time(), " : Successfully imported ", 
                  pathAndFn, sep = "")
      message(msg)
      msgV = c(msgV, msg)
    }
    else {
      msg = paste(Sys.time(), " : Failed to import ", pathAndFn, 
                  sep = "")
      warning(msg)
      msgV = c(msgV, msg)
      return(list(msg = msgV))
    }
    importedLineObj = EVFile[["Lines"]]$Item(nbr.of.lines.pre.import)
    if (is.null(lineName)) {
      msg = paste(Sys.time(), " : Imported line object name = ", 
                  importedLineObj$Name(), sep = "")
      message(msg)
      msgV = c(msgV, msg)
      invisible(list(lineObj = importedLineObj, msg = msgV))
    }
    if(!is.null(lineName))
    {
      line.check <- EVFile[["Lines"]]$FindByName(lineName)
      if (is.null(line.check)){
        msg = paste(Sys.time(), " : Renaming imported line object name from ", 
                    importedLineObj$Name(), " to ", lineName, sep = "")
        message(msg)
        msgV = c(msgV, msg)
        renameFlag=EVRenameLine(EVFile=EVFile,evLine=importedLineObj,newName=lineName)
        if(renameFlag){
          invisible(list(lineObj=importedLineObj,msg=msgV))
        } else {
          msg=paste(Sys.time(),' : failed to rename line.')
          warning(msg)
          msgV=c(msgV,msg)
          return(list(msg=msgV))
        }
      } else { #end of what to do if there isn't a line with the name of lineName
        #lineName existing:
        if(!overwriteExistingLine){
          msg=paste(Sys.time(),' : line with name',lineName,'already exists and will not be overwritten')
          warning(msg)
          msgV=c(msgV,msg)
          msg=paste(Sys.time(),' : created line object called',importedLineObj$Name(), 'is available as a COM object in $lineObj')
          message(msg)
          msgV=c(msgV,msg)
          return(list(lineObj=importedLineObj,msg=msgV))
        }
        if(overwriteExistingLine)
        {
          msg=paste(Sys.time(),' : attempting to overwrite existing line object called',lineName,'with line object called',importedLineObj$Name())
          message(msg)
          msgV=c(msgV,msg)
          overwriteFlag=line.check$overwriteWith(importedLineObj)
          if(overwriteFlag)
          {
            msg=paste(Sys.time(),' : exisiting line called',lineName,'was overwritten by imported EVL file')
            message(msg)
            msgV=c(msgV,msg)
            msg=paste(Sys.time(),': attempting to delete line called',importedLineObj$Name())
            message(msg)
            msgV=c(msgV,msg)
            EVDeleteLine(EVFile, EVFindLineByName(EVFile, importedLineObj$Name()))
            invisible(list(lineObj=line.check,msg=msgV))
          } else {
            msg=paste(Sys.time(),' : Failed to overwrite exisiting line called',lineName,' with imported EVL file')
            warning(msg)
            msgV=c(msgV,msg)
            return(list(lineObj=importedLineObj,msg=msgV))
            
          }  
        } #end of line overwrite exisiting line
      }#end of what to do if there is an existing line of name lineName
      
    } #end of what to do if lineName specified  
    
  } #end of import EVImportLine()  
