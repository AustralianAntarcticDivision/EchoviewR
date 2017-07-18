#' Rename an Echoview acoustic variable
#'
#'This function renames an existing acoustic variable
#'
#'@note The function will check if the proposed new variable new already exists.  If the porposed new name exists, then the new name will be chnaged to 'newName duplicate name'.
#' @param EVFile Echoview file COM object
#' @param acoVarName The acoustic variable name to which a new variable will be added.
#' @param newName New name of the acoustic variable
#' @return list with three elements: $varObj acoustic variable COM object; $objName new name of the variable object;$msg message vector 
#' @export
#'
EVRenameAcousticVar=function(EVFile,acoVarName,newName){
  msgV=paste(Sys.time(),'Attempted to change the name of',acoVarName,'to ',newName)
  message(msgV)
  AObj=EVAcoVarNameFinder(EVFile=EVFile, acoVarName)
  msgV=c(msgV,AObj$msg)
  AObj=AObj$EVVar
  msg=paste(Sys.time(),'Checking that the variable name',newName,'does not already exist')
  nVar=(EVFile[['Variables']]$Count())
  varNameVec=vector(mode='character',length=nVar)
  for(i in 1:nVar) varNameVec[i]=EVFile[['Variables']]$Item(i-1)$Name()
  
  if(newName %in% varNameVec)
  {
    
    msg=paste(Sys.time(),newName, 'already exists so renaming variable',acoVarName,'to',newName,'duplicate name')
    warning(msg)
    msgV=c(msgV,msg)
    newName=paste(newName,'duplicate name')
    } else {
    msg=paste(Sys.time(), 'Changing acoustic variable name from ',acoVarName, 'to','newName')
  }
  #Check EV names
  if(newName %in% c(AObj[['Name']],AObj[['FullName']]))
  {
    msg=paste(Sys.time(),'Within variable duplicate name found.' , newName,' is identical to one or both of the acoustic variables Name or FullName properties.')
    warning(msg)
    msgV=c(msgV,msg)
    msg=paste(Sys.time(),'Changing',newName,'to', newName,'dup')
    warning(msg)
    msgV=c(msgV,msg)
    newName=paste(newName,'dup')
  }  
  AObj[['ShortName']] =newName
  if(AObj$Name()==newName){
      msg=paste(Sys.time(),'Changed variable name from',acoVarName,'to',newName)
      message(msg)
      msgV=c(msgV,msg)
    }
  else
  {
    msg=paste(Sys.time(),'Failed to change variable name from',acoVarName,'to',newName)
    warning(msg)
    msgV=c(msgV,msg)
    }
  invisible(list(varObj=AObj,objName=AObj$Name(),msg=msgV))
}

