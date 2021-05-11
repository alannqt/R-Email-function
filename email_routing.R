#clear all memory
rm(list=ls())

#set the libraries needed
library(RDCOMClient)
#if you do not have RDCOMClient installed, uncomment below to install devtools and from github
#devtools::install_github("alannqt/RDCOMClient")
library(stringr)

#set working directory
setwd("C:/Users/alann/Documents/GitHub/R_Email_function")
attachment_path <- "C:/Users/alann/Documents/GitHub/R_Email_function/Attachment"

#initialise the app
outApp <- COMCreate("Outlook.Application")
outlooknamespace <- outApp$GetNameSpace("MAPI")

#set the foldername to use. The folder name will follow the folder name in your outlook email
sourcefolder <- "Archive"
outputfolder <- "Bank"

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# This portion is important as the function will try to point to the right address in your outlook mail to refer to that folder of interest, if you do not know how it works. 
# change nothing here.
#set email poointer function
setEmail <- function(foldersrc){
  assigned <- FALSE
  i <- 1
  while (assigned == FALSE){
    readsourcecheck <- tryCatch({readsource<- outlooknamespace$Folders(i)$Folders(foldersrc)},error = function(e){return(0)})
    if (is.numeric(readsourcecheck)){
      i <- i+1
      next
    }
    assigned <- TRUE
    return(outlooknamespace$Folders(i)$Folders(foldersrc))
  }
}

#assigned the right pointer to email
emails <- setEmail(sourcefolder)$Items
processedfolder <- setEmail(outputfolder)
#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

#sample of reading the content of email 1 (Put an email with attachment to see the output)
#to read content of all email in a folder use a while loop, the example shown is using a single email
# while (emails()$Count()>0){ }
emailsubject <- emails(1)$Subject()
emailbody <- emails(1)$Body() 
emailattachment <- emails(1)$Attachments()

if (emailattachment$Count() > 0){
  for (j in c(1:emailattachment$Count())){
    emails(1)$Attachments(j)$SaveAsFile(paste(attachment_path,emailattachment$Item(j)[['DisplayName']],sep ="/"))
  }
}

#moving email from source folder to Test folder
emails(1)$Move(processedfolder)

#forward email
forwardmail <- email(1)$Forward()
forwardmail[["To"]] <- "someone@emailaddress.com"

#uncomment this portion to send
#forwardmail$Send()

#reply email
replymail <- email(1)$Reply()
replymail[["To"]] <- "someone@emailaddress.com"

#uncomment this portion to send
#replymail$Send()

#To create a new email and send
outMail <- outApp$CreateItem(0)

outMail[["To"]] <- "someone@emailaddress.com"
outMail[["CC"]] <- "someone@emailaddress.com"
outMail[["Subject"]] <- "This is a test subject"
outMail[["Body"]] <- "Hi, this is a sample email content"
#uncomment this portion to send
#outMail$Send()

