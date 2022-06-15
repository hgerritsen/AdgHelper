# AdgHelper
Tools to help QC of ICES advice sheets

ICES produces catch advice for a large number of species on an annual basis (https://www.ices.dk/advice/Pages/Latest-Advice.aspx). The AdgHelper repository is intended to host scripts to help stock reviewers and auditors identify mistakes in the draft advice sheets.

This repository contains two R scripts and one MS word macro-enabled document:
* "copy advice sheets from sharepoint.R" does just that. It is a bid fiddly to get the url that lists the file names on the advice sharepoint but if you follow the instructions it should work. The alterative is to manually copy the files to a local folder.

* "accept changes macro.docm" lets you accept all changes for documents with tracked changes (the R function that reads the docx files does not correctly deal with tracked changes). You can do an entire batch of files in one go.

* "plot catchoptions.R" reads the catch options table from the MS Word advice sheets and creates plots to help identify any values that do not follow the same trend from the other data points. The same script also produces a csv file with all the tables in the Word document that can be used to further check the consistency of these tables (manually, in Excel)

