#-----Function for creating a word document with a stock assessment ------

#notes: This script updates the Stock Assessment section of the assessment document

#       any bit of text in the template that needs to be replaced
#       has to be defined as a bookmark in the word template. So, in the .docx file, 
#       highlight text, go to Insert/Bookmark, and give it a name. Make sure it's the same 
#       name used in the script. Also, make sure that sections that are bookmarked are one
#       single paragraph, i.e. no breaks, otherwise it won't overwrite the entire section

#       Text.list is a list with text corresponding to each section and names corresponding to
#       the bookmark name in the template

#       Figures.list is a list with path for the figures to be embedded in the report with names 
#       corresponding to the figure bookmark name in the template

#       Once finalised inserting Figure, remove all bookmarks in the report to remove line next 
#       to Figures

library(ReporteRs)

Possible.Data.types=c("CAES","Logbook","Processor returns","VMS","Recreational fishing survey",
                      "Charter returns","Economic","Environmental","Fishery-dependent",
                      "Fishery-independent survey","Tagging","Volunteer logbooks") 


Create.assessment=function(WDir,doc.filename,Template.loc,Text.ls,all_book,Fig.ls,Fig.w,
                           Fig.h,Fig.cap.ls,EQ.lst)
{
  setwd(WDir)
  #Create the word document 
  doc = docx(title=doc.filename,template=Template.loc) 
 
  #Update text
  Text.ls=Text.ls[!is.na(Text.ls)]
  BkMrk.txt=names(Text.ls)  
  for(i in 1:length(BkMrk.txt)) 
  {
    a=gsub("[\r\n]", "", Text.ls[[i]])
    a=gsub("    ", "", a)
    doc <- addParagraph(doc,value=a,bookmark=BkMrk.txt[i])
  }
    
    
    
  #Update figures
  BkMrk.fig=names(Fig.ls)
  BkMrk.fig.cap=names(Fig.cap.ls)  
  for(i in 1:length(BkMrk.fig))
  {
    doc=addImage(doc,Fig.ls[[i]],bookmark=BkMrk.fig[i],width=Fig.w[i],height=Fig.h[i])
    doc <- addParagraph(doc,value=Fig.cap.ls[[i]],bookmark=BkMrk.fig.cap[i])
  }
  

  doc = addMarkdown(doc, text = EQ.lst[[1]],
                       default.par.properties = parProperties(text.align = "justify",
                                                              padding.left = 0) )  
  #Update tables

  
  #delete un-used bookmarks
  Dop.book.mrk=all_book[-which(all_book%in%c(BkMrk.txt,BkMrk.fig,BkMrk.fig.cap))]
  if(length(Dop.book.mrk)>0)sapply(Dop.book.mrk,function(x) deleteBookmark(doc,x))
  
  #export document
  writeDoc( doc, file = doc.filename )
}


#This part goes in assessment.R#########################

#Create list of text
Text.list=list(
    Assessment_Overview="The different methods used by the Department to assess the status of 
    aquatic resources in WA have been categorised into five broad levels, ranging from relatively
    simple analysis ",
    Peer_review_of_Assessment="describe if/where peer reviewd",
    Data_used_in_assessment=paste(c("The information used in the assessment includes ",paste(
        Possible.Data.types[c(1,2,9,11)], "data",collapse=", "),". Write here something else."),collapse=""),
    Catch_by_sector=NA,
    Effort_by_sector=NA,
    Fishery_dependent_CPUE=NA,
    Fishery_independent_survey=NA,
    Catch_predictions=NA,
    Empirical_stock_recruitment=NA,
    Trends_in_size_age_composition=NA,
    
    PSA_productivity=NA,
    PSA_susceptibility=NA,
    PSA_overall_score=NA,
    PSA_limitations=NA,
    
    Biomass_dynamics_model_overview=NA,
    Biomass_dynamics_model_model_description=NA,
    Biomass_dynamics_model_inputs=NA,
    Biomass_dynamics_model_outputs=NA,
    Biomass_dynamics_model_uncertainty=NA,
    Biomass_dynamics_model_conclusion=NA,    
    
    Catch_curve_overview=NA,
    Catch_curve_model_description=NA,
    Catch_curve_model_inputs=NA,
    Catch_curve_model_outputs=NA,
    Catch_curve_model_uncertainty=NA,
    Catch_curve_model_conclusion=NA,    
    
    Per_recruit_analysis_overview=NA,
    Per_recruit_analysis_model_description=NA,
    Per_recruit_analysis_inputs=NA,
    Per_recruit_analysis_outputs=NA,
    Per_recruit_analysis_uncertainty=NA,
    Per_recruit_analysis_conclusion=NA,    
    
    Demography_overview=NA,
    Demography_model_description=NA,
    Demography_inputs=NA,
    Demography_outputs=NA,
    Demography_uncertainty=NA,
    Demography_conclusion=NA,    
    
    Age_size_structured_model_overview=NA,
    Age_size_structured_model_description=NA,
    Age_size_structured_model_inputs=NA,
    Age_size_structured_model_outputs=NA,
    Age_size_structured_model_uncertainty=NA,
    Age_size_structured_model_conclusion=NA,    
    
    Bio_economic_model_overview=NA,
    Bio_economic_model_description=NA,
    Bio_economic_model_inputs=NA,
    Bio_economic_model_outputs=NA,
    Bio_economic_model_uncertainty=NA,
    Bio_economic_model_conclusion=NA, 
    Other_modelling_methods=NA,
    
    Previous_assessment=NA,
    
    Weight_of_evidence_text=NA,
    Weight_of_evidence_conclusion=NA,
    Assessment_conclusion_advice=NA,
    Monitoring_implications=NA    
  )

#Create list of figures 
n.sp=1:2

  #figure location
if(!exists('handl_OneDrive')) source('C:/Users/myb/OneDrive - Department of Primary Industries and Regional Development/Matias/Analyses/SOURCE_SCRIPTS/Git_other/handl_OneDrive.R')

hndl_sp1=handl_OneDrive("Analyses/Population dynamics/Whiskery shark/2015/1_Outputs/")
hndl_sp2=handl_OneDrive("Analyses/Population dynamics/Gummy shark/2015/1_Outputs/")


Figures.list=list(
    paste(hndl_sp1,"Visualise data/avail.dat.tiff",sep=""),
    paste(hndl_sp2,"Visualise data/avail.dat.tiff",sep=""),
    paste(hndl_sp1,"Visualise data/Whiskery.catch.tiff",sep=""),
    paste(hndl_sp2,"Visualise data/Gummy.catch.tiff",sep="")
    
  )

names(Figures.list)=c(
    paste("Data_used_sp",n.sp,sep=""),
    paste("Catch_sp",n.sp,sep="")
  )

  #figure dimensions
Figures.width=c(6,6,6,6)
Figures.height=c(3,3,6,6)

  #figure captions
Figure.captions.list=list(
    "Figure 1. Data used species 1",
    "Figure 2. Data used species 2",
    "Figure 3. Catch species 1",
    "Figure 4. Catch species 2"
  )
names(Figure.captions.list)=c(
      paste("Caption_data_used_sp",n.sp,sep=""),
      paste("Caption_catch_sp",n.sp,sep="")
  )


#Create list of tables


#Create list of equations
Eqtn.list=list(mkd)

#All figure and table bookmarks
ALL_bkmrk=c(names(Text.list),names(Figures.list),names(Figure.captions.list))


#Run function with specified arguments
hndl=handl_OneDrive("Admin/Templates/")
Create.assessment(
    WDir=handl_OneDrive("Reports/Gummy_whiskery_DOF_assessment"),
    doc.filename="Assessment_2016.docx",
    Template.loc=paste(hndl,'Resource Report Template v.23 Feb 2016_Assessment.docx',sep=""),
    Text.ls=Text.list,
    all_book=ALL_bkmrk,
    Fig.ls=Figures.list,
    Fig.w=Figures.width,
    Fig.h=Figures.height,
    Fig.cap.ls=Figure.captions.list,
    EQ.lst=Eqtn.list
  )

