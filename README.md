# AdgHelper
Tools to help QC of ICES advice sheets

ICES produces catch advice for a large number of species on an annual basis (https://www.ices.dk/advice/Pages/Latest-Advice.aspx). The AdgHelper repository is intended to host  scripts to help stock reviewers and auditors identify mistakes.

For now there is just one R script. It reads the catch options table from the draft advice MS Word documents and creates plots to help identify any values that do not follow the same trend from the other data points.

I havent figured out how to automatically download all files from the advice sharepoint; because of the way the sharepoint site is set up only the first 30 files are displayed on the first page and i havent found a way to get the file locations of documents other than the first 30. So for now, you need to manually copy the word documents to a local folder. 

The second problem is that the R function does not correctly deal with tracked changes. The repository contains a macro-enabled Word file that will let you accept all changes for a batch of files to deal with this.
