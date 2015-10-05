# Prune-erickson Command Line Utility #

Below is documentation related to the Prune-Erickson Command line utility, an executable script that removes erroneous and sensitive data from excel files stored on your local computer. 

This repository is for the development files related to the utility. the readME is documentation for use of the app. The current version is 0.2.0 which is still under active development. 

Prune-erickson is a command line utility that has a very specific use for Erickson Living Employees only. 

This software is built off of the ```roo``` and ```Spreadsheet``` ruby libraries, which reads .xls and .xlsx file formats. 
Prune-erickson works by parsing excel file data, then copying necessary data from the spreadsheet and writing a new file that will contain necessary data. the original file is untouched, but an output file is created in the same folder which virtually has all unnecessary data removed. 

.XLS .XLSX  and .CSV formats are all supported. 

### Set up###

Requirements: for proper installation, it is important to be running on a MAC OSX system with the latest ruby release installed ```2.2.0```

you can begin by installing the app using the gem command on the command line

```$ gem install prune-erickson```

this will install the latest and most stable version, along with its dependencies. 

the general pattern of using this utility is to call the name of the utility and the arguments

```$prune-erickson [args]```

###helpful commands###

```help```
will give you a friendly reminder of how to process excel spreadsheets

```-v, -version, or version```
will give you the version of the utility

```$prune-erickson input-file.xls 1 output.xls```

this is the processing command, you must provide an input file which can be in the .xls, .xlsx, and .csv formats. the second argument is the sheet number, or the sheet where you need to remove data from your excel sheet. the last argument is the output file which only outputs files in the .xls format. 


### Who do I talk to? ###

* adlondono owns this repo and wrote this program
* feel free to make a pull request