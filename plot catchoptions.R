library(docxtractr)
library(dplyr)
library(tidyr)
library(ggplot2)


# need to copy (or sync) files from sharepoint and point to the relevant folder
# the function does not deal with track changes properly, so use word macro 
# to accept all changes in the docs: 'accept changes macro.docm'

# make sure no word files are open when you run this

#files <- list.files('C:/Users/hgerritsen/sharepoint','.docx',full.names=T,recursive=T)
files <- list.files('./2022','.docx',full.names=T,recursive=T)

outdir <- './plots'
dir.create(outdir)


for(f in files){
  cat(f,'\n')
  file <- rev(strsplit(f,'/')[[1]])[1]
  
  doc <- read_docx(f,track_changes = 'accept')
# track changes = 'accept' doesnt always work (even with pandoc installed)
# need to accept all changes in actual document first

  # plot the catch options table (if it exists)
  if(length(doc$tbls)<2) next() # if the document doesnt have at least 2 tables, skip
  tab2 <- docx_extract_tbl(doc,2) # extract 2nd table in doc
  print(tab2)
  cat('\n\n')
  if(grepl('Basis',names(tab2)[1],F)) {
    tab2 %>% mutate(across(-1, function(x) as.numeric(gsub("[a-zA-Z*^% ]","",x)) )) -> tab2a
    tab2a
    tab2b <- tab2a %>% pivot_longer(-(1:2)) %>% arrange(2)
    ggplot(subset(tab2b,!is.na(value)),aes_string(names(tab2b)[2],'value',col=names(tab2b)[1])) + 
      geom_line(aes(group=NA),col=1) + geom_point() +
      facet_wrap(~name,scales='free_y') + ggtitle(f)
    ggsave(file.path(outdir,gsub('[.]docx','.png',file)),width=6,height=4.5,dpi=600,scale=1.5)
  }
  
  # save all the tables in one csv doc, to explore in excel
  out <- NULL
  for(i in 1:length(doc$tbls)){
    tabi <- docx_extract_tbl(doc,i,header=F) # summary table is usually the last table
    tempfun <- function(x) gsub("[*^% ,]","",x) # remove footnote references and spaces
    tabi <- tabi %>% mutate(across(.fns=tempfun)) 
    
    out <- bind_rows(out,tabi,data.frame(V1=NA))
  }
  write.table(out,file.path(outdir,gsub('[.]docx','.csv',file)),row.names=F,col.names=F,sep=',',quote=F,na='')
}