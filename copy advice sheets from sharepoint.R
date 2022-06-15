library(httr)
library(XML)

#' A function to copy all word files from the advice sharepoint
#' 
#' @param user An ICES username with access to the sharepoint
#' @param password An ICES username with access to the sharepoint
#' @param site_url URL - see details
#' @param dir_out Location to save the word documents
#' 
#' @details 
#' To obtain the url that has a list of the files on the SharePoint, follow these steps:
#' In ICES Sharepoint browse to advice  2022 > Celtic Seas advice > draft advice
#' then select the library ribbon and click 'export to excel'
#' a file called owssvr.iqy will download
#' open the file and copy the url below. 
#' Do the same for the released advice and for other regions

giveMeSomeAdvice <- function(user,password,site_url,dir_out) {
  
  # extract the file names and paths
  response <- GET(site_url,authenticate(user,password, "ntlm"))
  content <- rawToChar(response$content)
  xmllist <- xmlToList(content)
  files <- unlist(lapply(xmllist$data,function(x) x['ows_LinkFilename']))
  folder <- gsub('%5f','_',gsub('%2f','/',rev(strsplit(site_url,'RootFolder=')[[1]])[1]))
  path <- paste0('https://community.ices.dk',folder)
  
  # copy all files to current working directory
  for(file in files){
    url <- paste0(path,'/',file)
    cat(url,'\n')
    file_out <- file.path(dir_out,file)
    GET(url,authenticate(user,password, "ntlm"),write_disk(file_out, overwrite = TRUE))
  }
}
  

#####

# need an ices username and password with access to the sharepoint
user <- 'gerritsen' 
password <- 'xxxx'

user <- 'whitej' 
password <- 'xxxx'

# this years advice
dir_out="F:/StockBooks/_Stockbook2022/ICES advice/2022"

# celtic seas draft
site_url <- "https://community.ices.dk/Advice/Advice2022/CelticSeas/_vti_bin/owssvr.dll?XMLDATA=1&List={3629B196-3FCE-460A-9FF0-D9F977993605}&View={02215852-72B6-420B-96E9-61B0CB45F09D}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fCelticSeas%2fDraft%5fadvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# celtic seas released
site_url <- "https://community.ices.dk/Advice/Advice2022/CelticSeas/_vti_bin/owssvr.dll?XMLDATA=1&List={5F00BF50-1090-4205-8679-6C8071420529}&View={2C21CB06-83A5-46F5-AFC8-E15DC0D7E204}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fCelticSeas%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# biscay draft
site_url <- "https://community.ices.dk/Advice/Advice2022/BayOfBiscay/_vti_bin/owssvr.dll?XMLDATA=1&List={4DD47079-60E3-49AF-BA61-FB6009E3C96E}&View={FE74E517-19CD-427E-B3F1-CD6610C7D6D7}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fBayOfBiscay%2fDraft%5fadvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# biscay released
site_url <- "https://community.ices.dk/Advice/Advice2022/BayOfBiscay/_vti_bin/owssvr.dll?XMLDATA=1&List={B26484C4-B089-401D-8612-91B0FD014812}&View={EC70901C-D8FE-4B5B-8158-DAA91622B280}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fBayOfBiscay%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# widely released
site_url <- "https://community.ices.dk/Advice/Advice2022/widely/_vti_bin/owssvr.dll?XMLDATA=1&List={1F30B4DC-FCE2-47F2-A795-05E1F979AC71}&View={C5737DF4-EA29-40B7-BA5D-13A1D0454A83}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fwidely%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# widely draft
site_url <- "https://community.ices.dk/Advice/Advice2022/widely/_vti_bin/owssvr.dll?XMLDATA=1&List={0A2DECC0-4FAB-48AB-A944-238DCE8B2FA2}&View={9A0BD319-C6DE-4242-82A7-260D7AE92B79}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fwidely%2fDraft%5fadvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# north sea released
site_url <- "https://community.ices.dk/Advice/Advice2022/northSea/_vti_bin/owssvr.dll?XMLDATA=1&List={3C0A2FC8-8D04-4AE9-8108-A4BE1A6D5686}&View={957DB8E5-483F-4BD9-990B-EE71DED463F6}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fnorthSea%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# north sea draft
site_url <- "https://community.ices.dk/Advice/Advice2022/northSea/_vti_bin/owssvr.dll?XMLDATA=1&List={9CD2E03D-C488-4322-A20F-D74A5EDFF2F9}&View={FEE381D5-3D7C-4852-9628-12CD65D92857}&RowLimit=0&RootFolder=%2fAdvice%2fAdvice2022%2fnorthSea%2fDraft%5fadvice"
giveMeSomeAdvice(user,password,site_url,dir_out)


# last year's
dir_out="F:/StockBooks/_Stockbook2022/ICES advice/2021"

site_url <- "https://community.ices.dk/Advice/2021/CelticSeas/_vti_bin/owssvr.dll?XMLDATA=1&List={451316AA-8313-44AC-B42A-0390369DCEDC}&View={9D551875-C972-4BDD-A5AF-D063C40C054D}&RowLimit=0&RootFolder=%2fAdvice%2f2021%2fCelticSeas%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

site_url <- "https://community.ices.dk/Advice/2021/BayOfBiscay/_vti_bin/owssvr.dll?XMLDATA=1&List={1A3E1B9B-8DF0-4DB6-AAE4-667E3FF81570}&View={D1B4B449-AE8F-48A9-B040-1AACF92CEFBB}&RowLimit=0&RootFolder=%2fAdvice%2f2021%2fBayOfBiscay%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

site_url <- "https://community.ices.dk/Advice/2021/Widely/_vti_bin/owssvr.dll?XMLDATA=1&List={5BE897B5-A7CC-439B-8404-F329A41C9A95}&View={340C7300-9B13-44B9-B973-6C0EFBB82FBC}&RowLimit=0&RootFolder=%2fAdvice%2f2021%2fWidely%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

site_url <- "https://community.ices.dk/Advice/2021/NorthSea/_vti_bin/owssvr.dll?XMLDATA=1&List={76AC9F13-6E3F-43DD-AA1B-B30290FC033D}&View={137007BE-CED9-4A84-91CA-56CB90C883D4}&RowLimit=0&RootFolder=%2fAdvice%2f2021%2fNorthSea%2fReleased%5fAdvice"
giveMeSomeAdvice(user,password,site_url,dir_out)

# now accept all changes using "accept changes macro.docm"