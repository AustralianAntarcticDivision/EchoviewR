#' Generate a coordinate list for a regular rectangular survey 
#'
#'The coordinate list is generated in degrees decimal degree format (dd.ddd), with 
#'Southern hemisphere denoted by negative numbers. Transect length and inter-transect 
#'spacing are specified in km and bearings in degress where North 0 deg, East 90 deg, 
#'South 180 deg and West 270 deg.
#'
#' @param startLon start longitude of survey.
#' @param startLat start latitude of survey.
#' @param lineLengthkm transect line length in km.
#' @param lineSpacingkm inter-transect spacing in km (see details).
#' @param startBearingdeg Orientation of survey grid in degrees.
#' @param numOfLines Number of transects.
#' @return Data frame of geographical coordinates of start (SOL) and end of line (EOL) positions
#' @export
#' @importFrom geosphere destPoint
#' @details Line spacing can be fixed, so specified by a single number or pseudo random (random draws from a uniform distribution), which is specified by a vector of two numbers, the minimun and maximum transect separation.
#' @seealso \code{\link{zigzagSurvey}}
#' @author Martin Cox \email{martin.cox@@aad.gov.au}
#' @examples
#' \dontrun{
#' (coords=lawnSurvey(startLon=-170,startLat=-60,lineLengthkm=2,lineSpacingkm=0.5,
#'startBearingdeg=30,numOfLines=5))
#'plot(0,0,xlim=range(coords[,1]),
#'ylim=range(coords[,2]),type='n',xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'legend('topright',c('Beginning','End'),col='blue',pch=c(17,15))
#'}
lawnSurvey=function(startLon,startLat,lineLengthkm,lineSpacingkm,startBearingdeg,numOfLines)
{
  if(length(lineSpacingkm)==1) lineSpacingkm=rep(lineSpacingkm,numOfLines-1)
  if(length(lineSpacingkm)==2) lineSpacingkm=runif(n=numOfLines-1,min=lineSpacingkm[1],max=lineSpacingkm[2])
  lineLengthm=lineLengthkm*1e3;lineSpacingm=lineSpacingkm*1e3
  direct=startBearingdeg+90
  direct[direct>360]=direct-360
  dir2=startBearingdeg+180
  dir2[dir2>360]=dir2-360
  lineBearings=c(startBearingdeg,dir2)
  out=matrix(NA,nrow=2*numOfLines,ncol=2,
             dimnames=list(paste(rep(c('SOL','EOL'),numOfLines),
                                 'line',sort(rep(1:numOfLines,2)),sep=''),
                           c('lon','lat')))
  out[1,]=c(startLon,startLat)
  out[2,]= geosphere::destPoint(p=out[1,],b=startBearingdeg,d=lineLengthm) 
  if(numOfLines>1){
    for(i in 1:(numOfLines-1))
    {  
      out[(i*2)+1,]=geosphere::destPoint(p=out[i*2,],b=direct,d=lineSpacingm[i])
      out[(i*2)+2,]=geosphere::destPoint(p=out[(i*2)+1,],b=lineBearings[ifelse(i%%2==1,2,1)],d=lineLengthm) 
    }
  }
  return(data.frame(out))
}


#' Generate a coordinate list for a zig-zag survey 
#'
#'The coordinate list is generated in degrees decimal degree format (dd.ddd), with 
#'Southern hemisphere denoted by negative numbers. Transect length are specified in km 
#'and bearings in degress where North 0 deg, East 90 deg, South 180 deg and West 270 deg.
#'
#' @param startLon start longitude of survey.
#' @param startLat start latitude of survey.
#' @param lineLengthkm transect line length in km.
#' @param startBearingdeg Bearing of each transect in degrees.
#' @param rotationdeg rotation angle for entire survey pattern.
#' @param numOfLines Number of transects.
#' @param proj4string projection string of class \link{CRS-class}
#' @param unrotated \code{FALSE} return rotated coordinates \code{TRUE} list of rotated and unrotated coordinates.
#' @return Geographical coordinates of start and end of line positions.  \code{unrotated=TRUE} list of rotated and unrotated coordinates
#' @export
#' @importFrom sp SpatialPoints CRS
#' @importFrom maptools elide
#' @importFrom stats na.omit
#' @importFrom utils capture.output read.csv read.table write.table
#' @seealso \link{zigzagSurvey}
#' @author Martin Cox \email{martin.cox@@aad.gov.au}
#' @examples
#' \dontrun{
#' coords=zigzagSurvey(startLon=-100,startLat=-60,lineLengthkm=2,
#' startBearingdeg=30,
#' rotationdeg=10,numOfLines=11)
#'plot(0,0,xlim=range(coords[,1]),ylim=range(coords[,2]),type='n',
#'xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'legend('topright',c('Beginning','End'),col='blue',pch=c(17,15))
#'#Use unrotated=TRUE and check coordinates in the- 
#'#-returned coordinate list object are identical:
#'coordL= zigzagSurvey(startLon=-100,startLat=-60,lineLengthkm=2,
#'startBearingdeg=30,
#' rotationdeg=0,numOfLines=11,unrotated=TRUE)
#' identical(coordL[[1]],coordL[[2]])
#' #display rotated and unrotated coordinates"
#'coordL= zigzagSurvey(startLon=-100,startLat=-60,lineLengthkm=2,
#'startBearingdeg=30,
#' rotationdeg=5,numOfLines=11,unrotated=TRUE) 
#' coords=coordL$unrotatedGeogs; coordsRotate=coordL$rotatedGeogs
#' plot(0,0,xlim=range(c(coords[,1],coordsRotate[,1])),
#' ylim=range(c(coords[,2],coordsRotate[,2])),type='n',
#' xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'arrows(x0=coordsRotate[1:(nrow(coordsRotate)-1),1], 
#'y0=coordsRotate[1:(nrow(coordsRotate)-1),2], 
#'       x1 = coordsRotate[2:nrow(coordsRotate),1], 
#'       y1 = coordsRotate[2:nrow(coordsRotate),2],col='blue')
#'legend('bottomleft',c('Unrotated','Rotated'),lty=1,col=c(1,'blue'))
#'}
zigzagSurvey=function(startLon,startLat,lineLengthkm,startBearingdeg,rotationdeg,numOfLines,
                      proj4string=CRS("+proj=longlat +datum=WGS84"),unrotated=FALSE)
{
  lineLengthm=lineLengthkm*1e3
  angs=c(startBearingdeg,180-startBearingdeg)
  lineIND=1:numOfLines
  out=matrix(NA,numOfLines+1,2,
             dimnames=list(c('SOLline1',paste('EOLline',lineIND,'SOLline',lineIND+1,sep='')[-numOfLines],paste('EOLline',numOfLines,sep='')),
                           c('lon','lat')))
  out[1,]=c(startLon,startLat)
  for(i in 1:numOfLines)
    out[i+1,]=destPoint(p=out[i,], b=angs[ifelse(i%%2==1,1,2)], d=lineLengthm)
  
  spout=SpatialPoints(out,proj4string=proj4string)
  rot=coordinates(elide(obj=spout,rotate=rotationdeg,center=coordinates(spout)[1,]))
  colnames(rot)=c('lon','lat')
  if(unrotated)
    return(list(unrotatedGeogs=out,rotatedGeogs=rot))
  return(rot)
}

#'Centre an zig-zag line transect survey on a given position
#'
#'Centres a zig-zag survey on a desired latitude and longitude
#'@param centreLon Desired centre location of survey
#'@param centreLat Desired centre location of survey
#'@param proj4string projection string of class \link{CRS-class}
#'@param tolerance maximum distance (in metres) between desired survey centre and realised survey centre
#'@param ... other arguments to be passed into \\link{zigzagSurvey}
#'@return Line transect coordinates as specified in \link{zigzagSurvey}
#'@details The call of \link{zigzagSurvey} has \code{unrotated=FALSE}
#'@importFrom geosphere geomean distHaversine
#'@export
#'@examples
#'\dontrun{
#'coords=centreZigZagOnPosition(centreLon=-33,centreLat=-57,lineLengthkm=60,startBearingdeg=30,
#'rotationdeg=10,numOfLines=21)
#'plot(0,0,xlim=range(coords[,1]),ylim=range(coords[,2]),
#'type='n',xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'points(-100,-60,col='purple',pch=19,cex=2)
#'points(geomean(coords),col='red',pch=19,cex=1)
#'legend('topright',c('Beginning','End','Desired centre','Actual centre'),
#'  col=c('blue','blue','purple','red'),pch=c(17,15,19,19),pt.cex=c(1,1,2,1))
#'  }
centreZigZagOnPosition=function(centreLon,centreLat,
                                proj4string=CRS("+proj=longlat +datum=WGS84"),tolerance=20,...)
{
  coords=zigzagSurvey(startLon=centreLon,startLat=centreLat,proj4string=proj4string,
                      unrotated=FALSE,...)
  centreCoords=(geomean(coords))
  shiftV=c(centreLon-centreCoords[1],centreLat-centreCoords[2])
  coords=coordinates(elide(obj=SpatialPoints(coords,proj4string=proj4string),shift=shiftV))
  colnames(coords)=c('lon','lat')
  shiftedCentreCoords=geomean(coords)
  errorD=distHaversine(p1=c(centreLon,centreLat), p2=geomean(coords))
  if(errorD>tolerance)
    stop(tolerance,' m tolerence between desired and realised survey centre exceeded.')
  message('Difference between desired and realised survey centre = ', round(errorD,2))
  return(coords)
}

#'Centre a regular rectangular survey on a given position
#'
#'Centres a regular rectangular survey on a desired latitude and longitude
#'@param centreLon Desired centre location of survey
#'@param centreLat Desired centre location of survey
#'@param proj4string projection string of class \link{CRS-class}.  If NULL defaults to "+proj=longlat +datum=WGS84"
#'@param tolerance maximum distance (in metres) between desired survey centre and realised survey centre
#'@param ... other arguments to be passed into \link{lawnSurvey}
#'@return Line transect coordinates (lon, lat) as specified in \link{lawnSurvey}
#'@importFrom geosphere geomean distHaversine
#'@importFrom maptools elide
#'@export
#'@examples
#'\dontrun{
#'coords=centreLawnOnPosition(centreLon=-170,centreLat=-60,lineLengthkm=2,lineSpacingkm=0.5,
#'startBearingdeg=30,numOfLines=5)
#'plot(0,0,xlim=range(coords[,1]),ylim=range(coords[,2]),type='n',
#'xlab='Longitude, deg',ylab='Latitude, deg')
#'arrows(x0=coords[1:(nrow(coords)-1),1], y0=coords[1:(nrow(coords)-1),2], 
#'       x1 = coords[2:nrow(coords),1], y1 = coords[2:nrow(coords),2])
#'text(coords,row.names(coords),cex=0.6)
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'points(-170,-60,col='purple',pch=19,cex=2)
#'points(geomean(coords),col='red',pch=19,cex=1)
#'legend('bottomright',c('Beginning','End','Desired centre','Actual centre'),
#'  col=c('blue','blue','purple','red'),pch=c(17,15,19,19),pt.cex=c(1,1,2,1))
#'}
centreLawnOnPosition=function(centreLon,centreLat,
                              proj4string=NULL,tolerance=20,...)
{
  
  if(is.null(proj4string))  proj4string=CRS("+proj=longlat +datum=WGS84")
  coords=lawnSurvey(startLon=centreLon,startLat=centreLat,...)
  centreCoords=(geomean(coords))
  shiftV=c(centreLon-centreCoords[1],centreLat-centreCoords[2])
  coords=coordinates(elide(obj=SpatialPoints(coords,proj4string=proj4string),shift=shiftV))
  colnames(coords)=c('lon','lat')
  shiftedCentreCoords=geomean(coords)
  errorD=distHaversine(p1=c(centreLon,centreLat), p2=geosphere::geomean(coords))
  if(errorD>tolerance)
    stop(tolerance,' m tolerence between desired and realised survey centre exceeded.')
  message('Difference between desired and realised survey centre = ', round(errorD,2),' m')
  return(coords)
}
