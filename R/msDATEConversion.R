#' Convert a Microsoft DATE object to a human readable date and time
#' 
#' Time stamps in Echoview, such as start and end times of, for example, invididual regions, use the Microsoft DATE format.  This function converts the Microsoft DATE object to a human readable date and time.  NB no time zone is returned.
#' @param dateObj a Microsoft date object
#' @return time stamp in the format yyyy-mm-dd hh:mm:ss
#' @seealso www.echoview.com
#' @export
msDATEConversion <- function (dateObj) {
  as.POSIXct(round(dateObj*60*60*24),origin="1899-12-30",tz="GMT")
}
