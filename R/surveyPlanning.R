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
#' @importFrom stats runif
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

#' Translate geographic coordinates to a target centre
#'
#' Shifts a set of geographic coordinates (longitude/latitude, decimal degrees) so that
#' their centroid matches a specified target location. The centring is performed in a
#' planar projected CRS (default CCAMLRGIS projection EPSG:6932, units in metres) by translating the point
#' cloud so its planar mean \eqn{(x,y)} equals the projected target centre. The shifted
#' coordinates are then transformed back to geographic coordinates (EPSG:4326, i.e. WGS84).
#' This function is useful for centring survey designs generated in geographic coordinates
#' (e.g. using \code{lawnSurvey} or \code{zigzagSurvey}) on a desired location \code{centreLon,centreLat}.
#' @param geogs A two-column matrix/data.frame of geographic coordinates with columns
#'   \code{lon} and \code{lat} (or equivalent order: longitude in column 1, latitude in column 2),
#'   in decimal degrees (EPSG:4326). 
#' @param centreLon Numeric. Target centre longitude (decimal degrees).
#' @param centreLat Numeric. Target centre latitude (decimal degrees).
#' @param planar_epsg Integer. EPSG code for the projected CRS used for the planar
#'   translation (metre units recommended). Default is \code{6932}.
#'
#' @return A data.frame with columns \code{lon} and \code{lat} containing the translated
#'   geographic coordinates (decimal degrees).
#'
#' @export
#'
#' @importFrom sf st_as_sf st_transform st_coordinates
#' @importFrom geosphere distGeo geomean
#'
#' @examples
#' \dontrun{
#' # Example input coordinates (lon/lat)
#' geogs <- data.frame(
#'   lon = c(-100, -99.9, -100.1),
#'   lat = c(-60, -60.05, -59.95)
#' )
#'
#' centred <- centre_geogs(geogs, centreLon = -100, centreLat = -60, planar_epsg = 6932)
#' }

centre_geogs<-function(geogs,centreLon,centreLat, planar_epsg= 6932)
{
  
  # Build sf points in lon/lat (EPSG:4326), then project to EPSG:6932 for rotation in metres
  pts_ll <- sf::st_as_sf(
    data.frame(lon = geogs[, 1], lat = geogs[, 2]),
    coords = c("lon", "lat"),
    crs = 4326,
    remove = FALSE
  )
  pts_xy <- sf::st_transform(pts_ll, crs=planar_epsg)
  xy <- sf::st_coordinates(pts_xy)[, 1:2, drop = FALSE]
  
  # project target centre geogs:
  centre_ll <- sf::st_as_sf(
    data.frame(lon = centreLon, lat = centreLat),
    coords = c("lon", "lat"),
    crs = 4326,
    remove = FALSE
  )
  centre_xy <- sf::st_transform(centre_ll, crs=planar_epsg)
  cent_xy <- sf::st_coordinates(centre_xy)[, 1:2, drop = FALSE]
  
  #planar translation:
  dX=cent_xy[1]-mean(xy[,1])
  dY=cent_xy[2]-mean(xy[,2])
  shiftedXY=xy+matrix(rep(c(dX,dY),nrow(xy)),ncol=2,byrow=TRUE)
  
  #translated geogs:
  # project target centre geogs:
  translated_xy <- sf::st_as_sf(
    data.frame(X = shiftedXY[,1], Y = shiftedXY[,2]),
    coords = c("X", "Y"),
    crs = planar_epsg,
    remove = FALSE
  )
  translated_ll <- sf::st_transform(translated_xy, crs=4326)
  trans_ll <- sf::st_coordinates(translated_ll)[, 1:2, drop = FALSE]
  trans_ll <- as.data.frame(trans_ll)
  names(trans_ll)=c('lon','lat')
  offset=geosphere::distGeo(cbind(centreLon,centreLat),
                            geosphere::geomean(trans_ll)) #m
  message(paste("Survey centred at (lon,lat)=(",round(centreLon,4),",",round(centreLat,4),
                ") with offset of ",round(offset,2)," m from required centre coordinates."))          
  return(trans_ll)
}









#' Generate a coordinate list for a zig-zag survey
#'
#' Generates a zig-zag survey track as a sequence of geographic
#' coordinates (longitude/latitude, decimal degrees). Transect lengths are specified
#' in kilometres and bearings in degrees, where North = 0°, East = 90° etc.
#'
#' Survey geometry is constructed in geographic coordinates (WGS84; EPSG:4326).
#' If a non-zero rotation is requested \code{rotationdeg}, the full survey pattern 
#' is projected to a planar coordinate reference system (default EPSG:6932, metres), 
#' rotated about the first point, and then transformed back to geographic coordinates for output.
#' NB \code{planar_epsg} can be used to define the planar coordinate system. 
#' 
#' @param startLon Numeric. Starting longitude of the survey (decimal degrees).
#' @param startLat Numeric. Starting latitude of the survey (decimal degrees).
#' @param lineLengthkm Numeric. Length of each transect line in kilometres.
#' @param startBearingdeg Numeric. Bearing of the first transect in degrees.
#' @param rotationdeg Numeric. Rotation angle (degrees) applied to the entire survey
#'   pattern. Positive values rotate counter-clockwise.
#' @param numOfLines Integer. Number of transects.
#' @param planar_epsg Integer. EPSG code of a projected coordinate reference system
#'   (units in metres) used internally for rotation. Default is 6932
#'   (Cylindrical Equal Area as used in \code{CCAMLRGIS}).
#' @param unrotated Logical. If \code{FALSE} (default), return only rotated coordinates.
#'   If \code{TRUE}, return a list containing both rotated and unrotated coordinates.
#'
#' @return If \code{unrotated = FALSE}, a numeric matrix with columns
#'   \code{lon} and \code{lat} giving the geographic coordinates of the start and end
#'   points of each transect.
#'
#'   If \code{unrotated = TRUE}, a list with elements:
#'   \describe{
#'     \item{unrotatedGeogs}{Matrix of unrotated geographic coordinates}
#'     \item{rotatedGeogs}{Matrix of rotated geographic coordinates}
#'   }
#'
#' @export
#'
#' @importFrom sf st_as_sf st_transform st_coordinates
#' @importFrom stats na.omit
#' @importFrom utils capture.output read.csv read.table write.table
#' @importFrom geosphere destPoint
#'
#' @seealso \link{zigzagSurvey}
#'
#' @author Martin Cox \email{martin.cox@@aad.gov.au}
#'
#' @examples
#' \dontrun{
#' coords <- zigzagSurvey(
#'   startLon = -100,
#'   startLat = -60,
#'   lineLengthkm = 2,
#'   startBearingdeg = 30,
#'   rotationdeg = 10,
#'   numOfLines = 11
#' )
#'
#' plot(0, 0,
#'      xlim = range(coords[,1]),
#'      ylim = range(coords[,2]),
#'      type = "n",
#'      xlab = "Longitude, deg",
#'      ylab = "Latitude, deg")
#'
#' arrows(coords[-nrow(coords),1], coords[-nrow(coords),2],
#'        coords[-1,1], coords[-1,2])
#'
#'points(coords[1,1],coords[1,2],col='blue',pch=17,cex=2)
#'points(coords[nrow(coords),1],coords[nrow(coords),2],col='blue',pch=15,cex=2)
#'legend('topleft',c('Beginning','End'),col='blue',pch=c(17,15))

#' coordL <- zigzagSurvey(
#'   startLon = -100,
#'   startLat = -60,
#'   lineLengthkm = 2,
#'   startBearingdeg = 30,
#'   rotationdeg = 10,
#'   numOfLines = 11,
#'   unrotated = TRUE
#' )
#'
#' plot(0, 0,
#'       xlim = range(coordL$rotatedGeogs[,1]),
#'      ylim = range(coordL$rotatedGeogs[,2]),
#'      type = "n",
#'      xlab = "Longitude, deg",
#'      ylab = "Latitude, deg")
#' arrows(coordL$rotatedGeogs[-nrow(coordL$rotatedGeogs),1], 
#'  coordL$rotatedGeogs[-nrow(coordL$rotatedGeogs),2],
#'        coordL$rotatedGeogs[-1,1], 
#'        coordL$rotatedGeogs[-1,2])
#' points(coordL$rotatedGeogs[1,1],
#'  coordL$rotatedGeogs[1,2],col='blue',pch=17,cex=2)
#' points(coordL$rotatedGeogs[nrow(coordL$rotatedGeogs),1],
#'  coordL$rotatedGeogs[nrow(coordL$rotatedGeogs),2],
#'  col='blue',pch=15,cex=2)
#' legend('topleft',c('Beginning','End'),col='blue',pch=c(17,15))
#' }
zigzagSurvey <- function(startLon, startLat, lineLengthkm, startBearingdeg, rotationdeg, numOfLines,
                         planar_epsg = 6932, unrotated = FALSE) {
  
  rotate_xy <- function(xy, angle_deg, center_xy) {
    a <- angle_deg * pi / 180
    R <- matrix(c(cos(a), -sin(a),
                  sin(a),  cos(a)), nrow = 2, byrow = TRUE)
    
    xy0 <- sweep(xy, 2, center_xy, "-")
    xy1 <- t(R %*% t(xy0))
    sweep(xy1, 2, center_xy, "+")
  }
  
  lineLengthm <- lineLengthkm * 1e3
  angs <- c(startBearingdeg, 180 - startBearingdeg)
  lineIND <- 1:numOfLines
  
  out <- matrix(
    NA_real_, numOfLines + 1, 2,
    dimnames = list(
      c(
        "SOLline1",
        paste("EOLline", lineIND, "SOLline", lineIND + 1, sep = "")[-numOfLines],
        paste("EOLline", numOfLines, sep = "")
      ),
      c("lon", "lat")
    )
  )
  
  out[1, ] <- c(startLon, startLat)
  
  for (i in 1:numOfLines) {
    out[i + 1, ] <- geosphere::destPoint(
      p = out[i, ],
      b = angs[ifelse(i %% 2 == 1, 1, 2)],
      d = lineLengthm
    )
  }
  
  # No rotation: return early (but still support unrotated=TRUE contract)
  if (isTRUE(all.equal(rotationdeg, 0))) {
    if (unrotated) return(list(unrotatedGeogs = out, rotatedGeogs = out))
    return(out)
  }
  
  # Build sf points in lon/lat (EPSG:4326), then project to EPSG:6932 for rotation in metres
  pts_ll <- sf::st_as_sf(
    data.frame(lon = out[, 1], lat = out[, 2]),
    coords = c("lon", "lat"),
    crs = 4326,
    remove = FALSE
  )
  
  pts_xy <- sf::st_transform(pts_ll, planar_epsg)
  xy <- sf::st_coordinates(pts_xy)[, 1:2, drop = FALSE]
  
  # Rotate about the first point, in projected (metre) space
  center_xy <- xy[1, ]
  xy_rot <- rotate_xy(xy, rotationdeg, center_xy)
  
  # Rebuild rotated geometry in EPSG:6932, then transform back to lon/lat for output
  pts_xy_rot <- sf::st_as_sf(
    data.frame(x = xy_rot[, 1], y = xy_rot[, 2]),
    coords = c("x", "y"),
    crs = planar_epsg
  )
  
  pts_ll_rot <- sf::st_transform(pts_xy_rot, 4326)
  rot_ll <- sf::st_coordinates(pts_ll_rot)[, 1:2, drop = FALSE]
  
  rot <- rot_ll
  colnames(rot) <- c("lon", "lat")
  rownames(rot) <- rownames(out)
  
  if (unrotated) return(list(unrotatedGeogs = out, rotatedGeogs = rot))
  rot
}


 