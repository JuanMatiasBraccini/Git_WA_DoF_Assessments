#--------------Generalized function for undertaking a weight of evidence approach -------------

#note:  assessment level (1-5) dependent on data availability

Weight.of.evidence=function(Level1,Level2,Level3,Level4,Level5)
{
  #Level 1 assessment
  #note: catch only
  
  Level1=NA
  
  #Level 2 assessment
  #note: Level 1 + fishery-dependent effort
  
  Level2=NA
  
  #Level 3 assessment
  #note: Level 1 and/or 2 + fishery dependent biological sampling of catch (e.g. average size,
  #       fishing mortality, etc, estimated from representative samples)
  
  Level3=NA
  
  #Level 4 assessment
  #note: Levels 1, 2 or 3 + either fishery-independed survey of relative abundance, exploitation rate,
  #       recruitment OR standardised fishery-dependent relative abundance
  
  Level4=NA
  
  #Level 5 assessment
  #note: Levels 1 to 3 and/or 4 integrated within a simulation stock assessment model
  
  Level5=NA
  
  #Qualitative risk assessment and Productivity Susceptibility Analysis
  #note: risk (within the next 5 years) = consequence X likelihood
  PSA=NA
  
  return(list(Level1=Level1,Level2=Level2,Level3=Level3,Level4=Level4,Level5=Level5,PSA=PSA))
}