#-----Function for creating a word document (in a smart way) ------
#copyright: Matias Braccini February 2016

#notes: The 'write.my.doc' function creates all sections of a document, including cover and
#       table of contents. It requires arguments with the working direction, headings, subheadings
#       body text, location of figures (which are imported as images), tables (imported as .csv files)
#       and reference lists. Most of these arguments have to be passed as data frames where an INDEX
#       is used to identified where in the text that information should be included.

#       For text, figures and tables, name each piece with the subheading extension to indicate
#       location in text

library(ReporteRs)
library(stringr)

write.my.doc=function(WDir,doc.filename,Doc.title,TOC,Hedin,Text,FIGS,TBL,Eqtns,FTRS,Txt.repos)
{
  #1. subfunctions
      #function for creating word table
  fn.word.table=function(TBL,HdR.col,HdR.bg,Hdr.fnt.sze,Hdr.bld,body.fnt.sze,Zebra,Zebra.col,Grid.col,Fnt.hdr,Fnt.body)
  {
    #create table
    MyFTable=FlexTable(TBL,add.rownames =F,
                       header.cell.props = cellProperties(background.color=HdR.bg), 
                       header.text.props = textProperties(color=HdR.col,font.size=Hdr.fnt.sze,
                                                          font.weight="bold",font.family =Fnt.hdr), 
                       body.text.props = textProperties(font.size=body.fnt.sze,font.family =Fnt.body))
    
    # zebra stripes - alternate colored backgrounds on table rows
    if(Zebra=="YES") MyFTable = setZebraStyle(MyFTable, odd = Zebra.col, even = "white" )
    
    # table borders
    MyFTable = setFlexTableBorders(MyFTable,
                                   inner.vertical = borderNone(),inner.horizontal = borderNone(),
                                   outer.vertical = borderNone(),
                                   outer.horizontal = borderProperties(color=Grid.col, style="solid", width=4))
    
    # set columns widths (in inches)
    #MyFTable = setFlexTableWidths( MyFTable, widths = Col.width)
    
    
    #add table to doc
    doc = addFlexTable(doc, MyFTable) 
  }
 
  #2. set working directory
  setwd(WDir)
  
  #3. Create the word document 
  doc = docx(title=doc.filename) 
  
  #4. Cover page
  if(sum(!is.na(Doc.title))>0)
  {
    for(s in 1:length(Doc.title)) doc<-addParagraph(doc,Doc.title[[s]],stylename="TitleDoc")
    doc <- addPageBreak(doc)    
  }
  
  #5. Add table of contents
  if(TOC=="YES")
  {
    doc <- addTitle( doc, "Table of contents", level =  1 )
    doc <- addTOC( doc)
    doc <- addPageBreak( doc )

    doc <- addTitle( doc, "List of figures", level =  1 )
    doc <- addTOC( doc, stylename = "Lgende" )
    doc <- addPageBreak( doc )
    
    doc <- addTitle( doc, "List of tables", level =  1 )
    doc <- addTOC( doc, stylename = "rTableLegend" )
    doc <- addPageBreak( doc )
    
  }
  
  #6. Order data
  Hedin$Index=as.character(Hedin$Index)
  Hedin$a=str_count(Hedin$Index, '_')
  Hedin$Level=Hedin$a+1
  Hedin$Nchr=nchar(Hedin$Index)
  Hedin$Ind1=as.numeric(with(Hedin,ifelse(a==0,Index,
                ifelse((a==1 & Nchr==3)|(a==2 & Nchr==5)|(a==3 & Nchr==7),substr(Index,1,1),
                ifelse((a==1 & Nchr==4)|(a==2 & Nchr%in%6:7)|(a==3 & Nchr%in%8:9),substr(Index,1,2),
                NA)))))
  Hedin$Ind2=as.numeric(with(Hedin,
                ifelse((a==1 & Nchr==3)|(a==2 & Nchr==5)|(a==3 & Nchr==7),substr(Index,3,3),
                ifelse((a==1 & Nchr==4)|(a==2 & Nchr%in%6:7)|(a==3 & Nchr%in%8:9),substr(Index,4,4),
                NA))))
  Hedin$Ind3=as.numeric(with(Hedin,
                ifelse((a==2 & Nchr==5)|(a==3 & Nchr==7),substr(Index,5,5),
                ifelse((a==2 & Nchr%in%6)|(a==3 & Nchr%in%8),substr(Index,6,6),
                ifelse((a==2 & Nchr%in%7)|(a==3 & Nchr%in%9),substr(Index,6,7),
                NA)))))
  Hedin$Ind4=as.numeric(with(Hedin,
                ifelse((a==3 & Nchr==7),substr(Index,7,7),
                ifelse((a==3 & Nchr%in%8),substr(Index,8,8),
                ifelse((a==3 & Nchr%in%9),substr(Index,9,9),
                NA)))))
  
  Hedin$Ind2[is.na(Hedin$Ind2)]=0
  Hedin$Ind3[is.na(Hedin$Ind3)]=0
  Hedin$Ind4[is.na(Hedin$Ind4)]=0
  
  Hedin=Hedin[order(Hedin$Ind1,Hedin$Ind2,Hedin$Ind3,Hedin$Ind4),]

  
  #7. Multi-loop for populating document
  INDX=Hedin$Index
  for(i in 1:length(INDX))
  {
    #add heading
    d=subset(Hedin,Index==INDX[i])
    doc <- addTitle( doc, as.character(d$Heading), level =  d$Level )
    
    #add text from script
    if(!is.null(Text))
    {
      Tx=subset(Text,Index==INDX[i])
      if(nrow(Tx)>0)
      {
        for(a in 1:nrow(Tx))
        {
          dat=Tx[a,]
          doc = addMarkdown(doc,text=dat$Text,
                            default.par.properties=parProperties(text.align="justify",padding.left=0))          
        }     
      }
    }
    
    #add text from repository
    Txt.repos$Location=as.character(Txt.repos$Location)
    txt.r=subset(Txt.repos,Index==INDX[i])
    if(nrow(txt.r)>0)
    {
      for(a in 1:nrow(txt.r))
      {
        dat=txt.r[a,]      
        doc <- addDocument( doc, filename = dat$Location )
      }      
    }
    
    #add figure
    FIGS$Lgnd=as.character(FIGS$Lgnd)
    FIGS$Location=as.character(FIGS$Location)
    Fig=subset(FIGS,Index==INDX[i])    
    if(nrow(Fig)>0)
    {
      for(a in 1:nrow(Fig))
      {
        dat=Fig[a,]      
        
        Captn=gsub("[\r\n]", "", dat$Lgnd)
        Captn=gsub("    ", "", Captn)
        doc=addImage(doc,dat$Location,width=dat$width,height=dat$height)   #figure 
        doc = addParagraph( doc, value=Captn, stylename = "Lgende" ) #caption below              
               
      }      
     }
    
    #add table
    TBL$Lgnd=as.character(TBL$Lgnd)
    TBL$Location=as.character(TBL$Location)
    tbl=subset(TBL,Index==INDX[i])    
    if(nrow(tbl)>0)
    {
      for(a in 1:nrow(tbl))
      {
        dat=tbl[a,]      
                      
        #caption
        Captn=gsub("[\r\n]", "", dat$Lgnd)
        Captn=gsub("    ", "", Captn)
        doc = addParagraph( doc, value=Captn, stylename = "Lgende" )
        #doc = addParagraph( doc, value=Captn, stylename = "rTableLegend",
        #        par.properties=parProperties( text.align="left",padding.left=0))
        
        #table
        TAB=read.csv(dat$Location)
        fn.word.table(TBL=TAB,HdR.col='black',HdR.bg='white',Hdr.fnt.sze=10,Hdr.bld='normal',
                      body.fnt.sze=10,Zebra='NO',Zebra.col='grey60',Grid.col='black',
                      Fnt.hdr= "Times New Roman",Fnt.body= "Times New Roman")
      }      
    }
    
    #add equations
    Eqtns$Location=as.character(Eqtns$Location)
    eq=subset(Eqtns,Index==INDX[i])
    if(nrow(eq)>0)
    {
      for(a in 1:nrow(eq))
      {
        dat=eq[a,]      
        doc <- addDocument( doc, filename = dat$Location )
      }      
    }
    
    #add special features
    FTRS$Location=as.character(FTRS$Location)
    ftrs=subset(FTRS,Index==INDX[i])
    if(nrow(ftrs)>0)
    {
      for(a in 1:nrow(ftrs))
      {
        dat=ftrs[a,]      
        doc <- addDocument( doc, filename = dat$Location )
      }      
    }
    
    
    
  }
  
  #8. Export document
  writeDoc( doc, file = doc.filename )
}