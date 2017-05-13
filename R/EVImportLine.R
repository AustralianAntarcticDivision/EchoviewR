#' Imports an Echoview Line file

#' This function imports an Echoview line file (.evl) using COM scripting
#' @param EVFile An Echoview file COM object
#' @param pathAndFn string path and filename to .evl file. 
#' @param lineName =NULL optional line name for imported line.
#' @return a list object with two elements- [[1]] imported EV line COM object, and [[2]] function message for log
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVCreateEditableLine}} 
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj, 'c:\\Users\\martin_cox\\Desktop\\KOASLine.EV')$EVFile
#' #testing...
#'  EVImportLine(EVFile)
#'  test two...
#' EVImportLine(EVFile,pathAndFn='c:\\Users\\martin_cox\\Desktop\\lineKAOS.evl')
#'
#' test three.. same line names (to be run after test2)
#' EVImportLine(EVFile,pathAndFn='c:\\Users\\martin_cox\\Desktop\\lineKAOS.evl',
#' lineName='Line6' )
#'
#' test four.. rename line with unique name
#' lineInfo=EVImportLine(EVFile,pathAndFn='c:\\Users\\martin_cox\\Desktop\\lineKAOS.evl',
#' lineName='seabed4' )
#' lineInfo  
#'}
EVImportLine <- function (EVFile, pathAndFn=NULL,lineName=NULL) {
  
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
    return(list(lineObj=importedLineObj,msg=msgV))
  }else{
    msg=paste(Sys.time(),' : Renaming imported line object name from ',importedLineObj$Name(),
              ' to ',lineName,sep='')
    message(msg)
    msgV=c(msgV,msg)
    #check if there is already a line with the new name
    line.check <- EVFile[["Lines"]]$FindByName(lineName)
    
    if (is.null(line.check) == FALSE) {
      msg=paste(Sys.time(), "Error: A line already exists with the name", lineName, ". Line not renamed", sep = " ")
      warning(msg)
      msgV=c(msgV,msg)
      return(list(lineObj=importedLineObj,msg=msgV))
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
  } #end of rename line if statement.
}
