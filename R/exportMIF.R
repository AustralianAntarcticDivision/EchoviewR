#'Write a map info file for import into Echoview.
#'
#'This function writes a polygon MIF file for import into Echoview and will
#'typically be used to export survey line transects from, for example, 
#'\link{lawnSurvey}.
#'@param coords a set of coordinates as a two column (longitude and latitude) matrix.
#'@param pathAndFileName character string MIF file export path and filename
#'@param pointNameExport Boolean (default=FALSE) export the name of each point.  
#'@param pointNameScaleFactor =c(0.005,0.03) see details
#'@details Exporting the name of each point uses the row name in the \code{coords} argument.  
#'
#'The \code{pointNameScaleFactor} argument is a two element numeric vector specifying the bounding 
#'box for each point name.  The bounding box is calculated for point name i by specifying the lower left hand corner 
#'of the point name position, with the upper right hand corner specified as 
#'c(lon_i,lat_i)+diff(range(lon_1..n,lat_1..n))*pointNameScaleFactor where n is the total number of points i.e nrow(coords)
#'@return Nothing
#'@export
#'@examples
#'\dontrun{
#'coords=lawnSurvey(startLon=-170,startLat=-60,lineLengthkm=2,lineSpacingkm=0.5,
#'startBearingdeg=30,numOfLines=5)
#'exportMIF(coords=coords,pathAndFileName='c:\\Users\\martin_cox\\Documents\\test4.mif')
#'}
exportMIF=function(coords,pathAndFileName,pointNameExport=FALSE,pointNameScaleFactor=c(0.05,0.03))
{
  textV=c('VERSION 300',
          'COLUMNS 0',
          'DATA',
          'PLINE',nrow(coords))
  for(i in 1:length(textV)) write.table(textV[i],pathAndFileName,quote=F,
                                        append=ifelse(i==1,F,T),row.names=F,
                                        col.names=F)
  
  if(pointNameExport){
    if(length(rownames(coords))!=0){
      dgeog=apply(coords,2,function(x)diff(range(x)))*pointNameScaleFactor
      for(i in 1:nrow(coords)){
        textV=paste('TEXT','\"',rownames(coords)[i],'\"')
        write.table(textV,pathAndFileName,quote=F,
                    append=T,row.names=F,col.names=F,sep='\t')
        write.table(cbind(coords[i,1],coords[i,2],coords[i,1]+dgeog[1],coords[i,2]+dgeog[1]),
                    pathAndFileName,quote=F,
                    append=T,row.names=F,col.names=F,sep='\t')
      }
    }
    else {stop('No row names found in coords ARG. Set pointNameExport=FALSE or add row names!')}
  }  
  write.table(coords,pathAndFileName,quote=F,
              append=T,row.names=F,col.names=F,sep='\t')
  return(paste('')) 
}
