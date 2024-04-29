#' Add a variable to an Echoview file
#' 
#' This function adds an Echoview variable to an existing Echoview file.  The variable is added by specifying the existing variable to which the new variable is added and the new variable type.
#' @details The new variable type is specifed using the variable-specific 'EOperator' integer.   For example, the 'Ping subset' operator 'EOperator' is 62. See Echoview help for details.  
#' @note This function is under development.  We hope to add an 'EOperator' look up table sp that the type of variable to add can be specified by its character name, rather than 'EOperator'   

#' @param EVFile Echoview file COM object
#' @param acoVarName The acoustic variable name to which a new variable will be added.
#' @param EOperator The 'EOperator' number of the variable to add. See notes
#' @param addedVarName =NULL. Name of the added variable.
#' @param ... additional arguments (up to 3) enable to assign parent variable to the added variable, e.g. single target detection
#' @return list with three elements: $varObj acoustic variable COM object; $objName new name of the variable object;$msg message vector 

#' @keywords Echoview COM scripting
#' @author Martin Cox, modified by Yang Yang
#' @seealso \link{EVAcoVarNameFinder}
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}

EVAddVariable <- function(EVFile, acoVarName, EOperator, addedVarName = NULL, ...) {
  msgV <- paste(Sys.time(), 'Attempted to add variable EOperator number to', acoVarName)
  message(msgV)
  
  AObj <- EVAcoVarNameFinder(EVFile = EVFile, acoVarName)
  msgV <- c(msgV, AObj$msg)
  AObj <- AObj$EVVar
  
  newAObj <- AObj$AddVariable(EOperator)
  
  cNewName <- newAObj$Name()
  
  if (class(newAObj)[1] != "COMIDispatch") {
    msg <- paste(Sys.time(), 'New acoustic variable not created')
    warning(msg)
    msgV <- c(msgV, msg)
    return(list(varObj = NULL, objName = NULL, msg = msgV))
  }
  
  if (!is.null(addedVarName)) {
    nnRes <- EVRenameAcousticVar(EVFile = EVFile, acoVarName = cNewName, newName = addedVarName)
    return(nnRes)
  } else {
    return(list(varObj = newAObj, objName = newAObj$Name(), msg = msgV))
  }
  
  # Process additional optional arguments
  if (length(list(...)) == 1) {
    acoVarName2 <- ...[[1]] # Extracting the first additional argument
    AObj2 <- EVAcoVarNameFinder(EVFile = EVFile, acoVarName2)
    AObj2 <- AObj2$EVVar
    newAObj$SetOperand(2, AObj2)
  }
  if (length(list(...)) == 2) {
    acoVarName3 <- ...[[2]] # Extracting the second additional argument
    AObj3 <- EVAcoVarNameFinder(EVFile = EVFile, acoVarName3)
    AObj3 <- AObj3$EVVar
    newAObj$SetOperand(3, AObj3)
  }
}



