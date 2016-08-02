#'The Jolly and Hampton (1990) estimator
#'
#'The Jolly and Hampton (1990) estimator is often used to calculate mean pelagic animal densities and associated variances.
#'
#'@param transectLength vector of parameter lengths
#'@param transectMeanDen transect weighted mean densities. See details.
#'@param transectName =NULL Optional character vector of transect names.
#'@param svyName =NULL Optional character variable of survey names.  Must be used if the \link{jhMultipleStrataF} function is used to estimate mean and variance across multiple strata.
#'@param area =NULL Optional character variable of the survey or stratum name.  Must be used if the \link{jhMultipleStrataF} function is used to estimate mean and variance across multiple strata.
#'@return two element list. 
#'@references Jolly, G. M., and Hampton, I. 1990. A stratified random transect design for acoustic surveys of fish stocks. Canadian Journal of Fisheries and Aquatic Sciences, 47: 1282-1291.
#'@seealso \link{jhMultipleStrataF}
#'@details The \code{transectMeanDen} In the case of acoustic data, each element (areal density value) is multiplied by the integration interval (along transect) length.  The distance multiplication must take place outside of this function.
#'@export
#'@examples
#'\dontrun{
#'tLengths=c(77.67,78.28,78.91,80.22,78.07,73.60,78.53,78.39,78.52,79.48)
#'denV=c(158.03,159.50,96.41,133.86,246.60,243.95,122.65,276.92,59.83,89.90)
#'svyA=80*100*1e6
#'jhF(transectLength=tLengths,
#'transectMeanDen=denV,
#'transectName=NULL,svyName='test',area=svyA)
#'}
jhF=function(transectLength,transectMeanDen,transectName=NULL,svyName=NULL,area=NULL)
{
  df=data.frame(transectLength=transectLength)
  meanTransectLength=mean(transectLength)
  df$transectLengthWt=df$transectLength/meanTransectLength
  df$transectDensity=transectMeanDen
  meanTransectDensity=mean(transectMeanDen)
  df$transectDensityWt=transectMeanDen*df$transectLengthWt
  meanTransectDensityWt=mean(df$transectDensityWt)
  df$densitydeviation=(transectMeanDen-meanTransectDensityWt)**2
  df$transWtsqTimesDD=df$transectLengthWt**2*df$densitydeviation
  rName=1:length(transectLength)
  if(!is.null(transectName)) rName=transectName
  row.names(df)=rName
  svy='Survey'
  if(!is.null(svyName)) svy=svyName
  stats=data.frame(surveyName=svy,MeanWeightedDensity=meanTransectDensityWt,
                   variance=sum(df$transWtsqTimesDD)/(nrow(df)*(nrow(df)-1)))
  stats$CV=sqrt(stats$variance)/stats$MeanWeightedDensity
  if(!is.null(area)){
    stats$area=area
    stats$BiomassTonnes=area*meanTransectDensityWt/1e6
    stats$lowerIntervalBiomassTonnes=area*(meanTransectDensityWt-1.96*sqrt(stats$variance))/1e6
    stats$upperIntervalBiomassTonnes=area*(meanTransectDensityWt+1.96*sqrt(stats$variance))/1e6}
  return(list(summaryStats=stats,transects=df))                 
}

#'Estimate mean and variance across multiple strata using the Jolly and Hampton (1990) method
#'
#'@param strata list of multiple strata, each element must be the result of a call of \link{jhF}.  See details.
#'@details The call of \link{jhF} used to create an element in the  \code{strata} argument must have the name of the stratum and the stratum area.
#'#'@references Jolly, G. M., and Hampton, I. 1990. A stratified random transect design for acoustic surveys of fish stocks. Canadian Journal of Fisheries and Aquatic Sciences, 47: 1282-1291.
#'@seealso \link{jhF}
#'@export
#'@examples
#'\dontrun{
#'svy1=jhF(transectLength=tLengths,
#'transectMeanDen=denV,
#'transectName=NULL,svyName='test',area=svyA)
#'svy2=jhF(transectLength=tLengths,
#'transectMeanDen=denV,
#'transectName=NULL,svyName='test',area=svyA)
#'
#'jhMultipleStrataF(strata=list(svy1,svy2))
#'
#'svy3=jhF(transectLength=tLengths,
#'transectMeanDen=denV,
#'transectName=NULL,svyName='test',area=svyA*2)
#'
#'jhMultipleStrataF(strata=list(svy1,svy3))
#'
#'svy4=jhF(transectLength=tLengths,
#'transectMeanDen=denV*2,
#'transectName=NULL,svyName='test',area=svyA)
#'
#'jhMultipleStrataF(strata=list(svy1,svy4))
#'
#'}
jhMultipleStrataF=function(strata)
{
  A=sapply(strata,function(x) x$summaryStats$area)
  means=sapply(strata,function(x) x$summaryStats$MeanWeightedDensity)
  vars=sapply(strata,function(x) x$summaryStats$variance)
  strataNames=sapply(strata,function(x) x$summaryStats$surveyName)
  overallMean=sum(A*means)/sum(A)
  overallVar=sum(vars*A**2)/sum(A)**2
  area=sum(A)
  strata=data.frame(strataMean=means,strataVar=vars,strataCV=sqrt(vars)/means)
  if(area>0)
    overallEstimates=data.frame(overallMean.g.m2=overallMean,
                                overallVar=overallVar,
                                overallCV=sqrt(overallVar)/overallMean,
                                overallArea.m2=area,
                                BiomassTonnes=area*overallMean/1e6,
                                lowerIntervalBiomassTonnes=area*(overallMean-1.96*sqrt(overallVar))/1e6,
                                upperIntervalBiomassTonnes=area*(overallMean+1.96*sqrt(overallVar))/1e6)
  if(is.null(area))
    overallEstimates=data.frame(overallMean=overallMean.g.m2,overallVar=overallVar,
                                overallCV=sqrt(overallVar)/overallMean,overallArea.m2=area)
  
  
  return(list(strata=strata,overallEstimates=overallEstimates))
}  
