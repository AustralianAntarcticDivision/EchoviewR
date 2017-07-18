#' Calculate the relative dimension of a school
#'
#'This function calculates the relative school dimension of Diner (2001)
#'
#'@param Sv Observed mean volume backscattering strength (Sv) of a school
#'@param L Observed (uncorrected) school length, m
#'@param P Mean depth of school 
#'@param Th Minimum Sv data threshold used to detect schools using the SHAPES algorithm
#'@param theta Transducer 3 dB beam width, degrees
#'@return Relative school dimension (Nb)
#'@references Diner, N. (2001). Correction on school geometry and density: approach based on acoustic image simulation. Aquatic Living Resources, 14(4), 211-222.
#'
#'@seealso \link{EVSchoolsDetect} \link{EVIntegrationByRegionsExport}
#'@export

#P=mean depth of the echotrace.

schoolRelativeDimension=function(Sv,L,P,Th,theta){
  
  dST=Sv-Th
  
  B=0.44*theta * dST^0.45#detection angle
  
  B=pi*B/180
  Nb=L/(2*P*tan(B/2))
  return(Nb)}
