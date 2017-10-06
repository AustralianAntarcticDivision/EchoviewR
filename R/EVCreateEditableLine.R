#' Create an EV Editable Line from an existing line
#'
#' This function creates  an EV Line in an EV file object by name using COM scripting.
#' @param EVFile An Echoview file COM object
#' @param lineNameToCopy a string containing the name of the line to copy
#' @param editableLineName =NULL desired editable line name.If null, EDIT is appended to lineNameToCopy and used as editable line name.
#' @param Multiply =1 multiply source line depth when creating new editable line
#' @param Add =0 add to source line depth when creating new editable line
#' @param SpanGaps =FALSE span gaps when creating new editable line
#' @return an Echoview editable line object
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVFindLineByName}} \code{\link{EVOpenFile}}
#' @examples
#' \dontrun{
#' EVAppObj <- COMCreate('EchoviewCom.EvApplication')
#' EVFile <- EVOpenFile(EVAppObj, '~~KAOS/KAOStemplate.EV')$EVFile
#' 
#' EVCreateEditableLine(EVFile = EVFile, 
#'     lineNameToCopy='Offset seabed',
#'     editableLineName='Editable Seabed line')
#'}
EVCreateEditableLine <- function (EVFile, lineNameToCopy, editableLineName = NULL, 
                                  Multiply = 1, Add = 0, SpanGaps = FALSE) {
  
  
  lineToCopy <- EVFindLineByName(EVFile=EVFile,lineName=lineNameToCopy)
  prenbrLines <- EVFile[["Lines"]]$Count()
  
  newEditLine <- EVFile[["Lines"]]$CreateOffsetLinear(lineToCopy, Multiply, Add,SpanGaps)
  
  if (EVFile[["Lines"]]$Count() == prenbrLines) {
    stop('Failed to create editable line')
  }
  if (is.null(editableLineName)) {
    editableLineName = paste('EDIT', lineNameToCopy, sep = '')
  }
  
  renameFlag <- EVRenameLine(EVFile, evLine = newEditLine, newName = editableLineName)
  
  if (!renameFlag) {
    msg <- paste(Sys.time(), ' Rename failed, line name already exisits, attempting to force line over write....')
    message(msg)
    lineToOverwrite=EVFindLineByName(EVFile=EVFile,lineName=editableLineName)
    msg <- paste(Sys.time(),'Over writing line name ',lineToOverwrite$Name(),'with',newEditLine$Name())
    message(msg)
    lineToOverwrite$OverwriteWith(newEditLine)
    msg <- paste(Sys.time(),' Deleting ',newEditLine$Name(),'...')
    EVFile[['Lines']]$Delete(newEditLine)
  } else {
    msg <- "Finished creating new line"
    lineToOverwrite = newEditLine
  }
  
  invisible(list(msg = msg, newEditLineObj = lineToOverwrite))
}
