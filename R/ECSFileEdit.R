#' A function to modify Ex60/70 Echoview calibration files (ECS files)
#' 
#' This function parses an EX60/70 Echoview calibration, finds a calibration parameter and changes its value.
#' 
#' @param ECSPathAndFile Echoview calibration path and filename (character)
#' @param ECStargetVar Calibration parameter to change (character)
#' @param NewValue Calibration values (numeric, vector)
#' @param maxNlines = 200. Maximum number of lines in ECS file to search (numeric).
#' @param nBlanks2Stop =4. Stop searching ECS file when this number of consecutive blanks has been encountered (numeric).
#' @param FileNameSuffix filename suffix used to generate modified ECS file (character).
#' @note This function is under development and has only been tested on ECS files generated in the 'SimradEK60Raw' format and that include FILESET and SOURCECAL settings only 
#' @note Currently the length of the NewValue parameter must equal the number of instances of the ECStargetVar found in the ECS file.
#' @details This function can be used to modify parameters in an Echoview calibration file.  The current form of the function is quite limited and we strongly recommend that users carefully check any resultant modified ECS files. Found instances of the target calibration parameter be retained by setting the corresponding element in the NewValue vector to NA.
#' @keywords Echoview COM scripting calibration ECS
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVAddCalibrationFile}}
#' @examples
#' \dontrun{
#' ECSFile<-'/20120326_KAOS_SimradEK5.ecs'
#' #change the absorption coefficient for the first frequency in 3-frequency EK60 data:
#' modifyECS(ECSPathAndFile=ECSFile,ECStargetVar='AbsorptionCoefficient',NewValue=c(1,NA,NA),
#' maxNlines=200,nBlanks2Stop=4,FileNameSuffix='alphaMod')
#' #change sound speed in FILESET and the three frequencies in SOURCECAL:
#' modifyECS(ECSPathAndFile=ECSFile,ECStargetVar='SoundSpeed',NewValue=rep(1500,4),
#'          maxNlines=200,nBlanks2Stop=4,FileNameSuffix='sosMod')
#' }
modifyECS=function(ECSPathAndFile,ECStargetVar,NewValue,maxNlines=200,nBlanks2Stop=4,FileNameSuffix)
{  
 #function to find lines with ECS text of interest:
  finder=function(ECSPathAndFile,maxNlines,ECSVar,verbose=FALSE){
    LV=vector(mode='numeric',length=maxNlines)
    locV=vector()
    for(i in 0:maxNlines){
    CL=scan(ECSPathAndFile,what=character(),skip=i,nlines=1,quiet=T)
    LV[i]=length(CL)
    grepTEST=grep(ECSVar,CL)
    if(length(grepTEST)>0)
      locV=append(locV,i)
    if(i>nBlanks2Stop){
      if(sum(LV[(i-nBlanks2Stop):i])==0){
        if(verbose)
          message(Sys.time(),' : ',nBlanks2Stop,' consecutive bank lines found in  ',
              ECSPathAndFile,'.  Stopped file reading at line number ', i)
      break
      }}
    }
    
    if(length(locV)<1) warning(Sys.time(), ' : No instances of calibration parameter',ECSVar,' found in ',ECSPathAndFile)
    if(length(locV)>0 & verbose) message(Sys.time(), ' : ',length(locV), ' instances of calibration parameter',ECSVar,' found in ',ECSPathAndFile)
  return(list(locV=locV,maxReadLine=i))
  }
  
  #function to find a calibration parameter value
  parameterValueF=function(ECSPathAndFile,loc) suppressWarnings(na.omit(as.numeric(scan(ECSPathAndFile,what=character(),skip=loc,nlines=1,quiet=T)))[1])
  settingNameV=c('FILESET','SOURCECAL','LOCALCAL')
  settingLocV=vector(mode='numeric',length=length(settingNameV))
  for(i in 1:length(settingLocV))
    settingLocV[i]=finder(ECSPathAndFile=ECSPathAndFile,maxNlines=maxNlines,
                          ECSVar=settingNameV[i])$locV
  maxLine=finder(ECSPathAndFile=ECSPathAndFile,maxNlines=maxNlines,
                 ECSVar=settingNameV[i])$maxReadLine-nBlanks2Stop
  lineDF=data.frame(setting=settingNameV,start=settingLocV,
                    stop=c(settingLocV[2:length(settingLocV)]-1,maxLine))
  
  
  freqLOC=finder(ECSPathAndFile=ECSPathAndFile,maxNlines=maxNlines,
                 ECSVar='Frequency')$locV
  
  if(length(freqLOC)<1) warning(Sys.time(), ' : No echosounder frequencies found in ',ECSPathAndFile)
  
  freqsetting=freqVal=vector(mode='character',length=length(freqLOC))
  for(i in 1:length(freqLOC)){
    IND=lineDF$start[i]<freqLOC[i] & lineDF$stop[i]>freqLOC[i]
    freqsetting[IND]=as.character(lineDF$setting[i])
    freqVal[i]=parameterValueF(ECSPathAndFile,freqLOC[i])}
  
  channel=tolower(freqsetting)
  channel=paste(toupper(substr(channel,1,1)),substr(channel,2,nchar(channel)-3),
                toupper(substr(channel,nchar(channel)-2,nchar(channel)-2)),
                substr(channel,nchar(channel)-1,nchar(channel)),sep='')
  channel=paste(channel,' T',1:length(channel),sep='')
  
  freqDF=data.frame(setting=freqsetting,channel=channel,
                    freq=as.numeric(freqVal))
  
  if(any(diff(freqDF$freq)<0))
    warning(Sys.time(),' : Check frequency and channel allocation: non ascending frequency to channel allocation found')
  
  message('--------------------------------')
  message(' CALIBRATION PARAMETER SUMMARY ')
  message(paste(capture.output(freqDF),collapse='\n'))
  message('--------------------------------')
  #find target calibration variable:
  locL=finder(ECSPathAndFile=ECSPathAndFile,maxNlines=maxNlines,
              ECSVar=ECStargetVar,verbose=TRUE)
  locV=locL[[1]];i=locL[[2]]

  ECSPathAndFile.out=gsub('.ecs',paste(FileNameSuffix,'.ecs',sep=''),ECSPathAndFile)
  
  #check for equal length of NewValues and found calibration parameters
  if(length(NewValue)!=length(locV))
    stop('Length of NewValue ARG must equal number of ',ECStargetVar,
         ' calibration parameters found in ECS file.','\n',  ' Currently, ',length(locV), ' instances of the calibration parameter ',
         ECStargetVar, ' were found \n  and the NewValue argument has ', length(NewValue),' elements.')
  
  if(any(is.na(NewValue)))
  {
    message(Sys.time(),' : Skipping calibration parameters with NA values in NewValue')
    locV=locV[!is.na(NewValue)]
    NewValue=NewValue[!is.na(NewValue)]
  }
  
  NewValCounter=1
  
  for(j in 0:(i-nBlanks2Stop)){


    CL=scan(ECSPathAndFile,what=character(),skip=j,nlines=1,quiet=T)
    #CL=gsub("﻿","",CL)
    CL=gsub('\\s+','',CL)
    if(j %in% locV){
      message(Sys.time(),' : Modifying calibration parameter, ',ECStargetVar,'.  Current value = ',parameterValueF(ECSPathAndFile,j))
      message(Sys.time(),' : Modifying calibration parameter, ',ECStargetVar,' with value ',NewValue[NewValCounter])
      
      CL[(1+which(CL=='='))]=paste(NewValue[NewValCounter],'#')
      if(CL[1]=='#'){
        message(Sys.time(),' : removing comment tag from calibration parameter')
      CL=CL[-1]}
      NewValCounter=NewValCounter+1

    }
    write.table(paste(CL,collapse=' '),ECSPathAndFile.out,append=ifelse(j==0,F,T),quote=F,sep=' ',col.names=F,row.names=F)
  }
}
  




