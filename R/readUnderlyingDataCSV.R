#'Read CSV resulting from a call of EVExportUnderlying()
#'
#'This function reads a CSV file created after a call of EVExportUnderlying().  The data can be Sv values or values from, for example, the Echoview formula operator.
#'
#'@param pathAndFn The path and filename of the CSV fileonlyHeader
#'@param samplePreFix ='sample' (character)text to add to the sample column names
#'@param depthPreFix = 'z' (character) text to add to the sample row names (depth)
#'@param asList = FALSE (boolean) return the CSV as a list. See details.
#'@param onlyHeader = FALSE (boolean) return only the header of the CSV, ignored if asList = TRUE. See details.
#'@param lastHeaderCol = 13 The last column of the header information.
#'@returns see details
#'@details The default return option is a data.frame of ping header and sample information.  
#'If asList=TRUE a list with $header and $samples is returned,
#'If onlyHeader=TRUE only the headerinformation is returned.  
#'@keywords Echoview COM scripting
#'@export
#'@references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#'@seealso \code{\link{EVExportUnderlying}}


readUnderlyingDataCSV=function(pathAndFn,samplePreFix='sample',depthPreFix='z',
                               asList=FALSE,onlyHeader=FALSE,lastHeaderCol=13)
{
  dat=read.csv(pathAndFn)
  tmpNames=names(dat)
  dat=cbind.data.frame(as.numeric(row.names(dat)),dat)
  digitWidth=nchar(as.character(ncol(dat)-lastHeaderCol))
  names(dat)=c(strsplit(tmpNames[1],'\\.\\.')[[1]][2],tmpNames[-1],
               paste(samplePreFix,formatC(1:(ncol(dat)-lastHeaderCol),flag='0',width=digitWidth),sep=''))
  if(asList) {
    samp=matrix(dat[,(lastHeaderCol+1):ncol(dat)],nrow=(ncol(dat)-lastHeaderCol),byrow=TRUE)
    row.names(samp)=paste(depthPreFix,seq(dat$Depth_start[1],dat$Depth_stop[1],length.out=dat$Sample_count[1]),sep='')
    return(list(header=dat[,1:lastHeaderCol],samples=samp))}
  if(onlyHeader) return(dat[,1:lastHeaderCol])
  return(dat)
}