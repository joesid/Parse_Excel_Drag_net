# Parse_Excel_Drag_net

Currently working as an IT Manage/IT Support Specialist 

I was asked to provide assistance with broadcasting a message with unique attachments and a personalized message, I normally use a "Mail Merge" Addon on Gmail for this,

Mail merge requires you use a Google Sheet with Colunms  like this e.g->     |First Name | Middle Name | Last Name | Email Address | File Attachment ID |

The data in the columns can be used to personalize emails e.g Dear {First Name} for each email sent, the name data will be taken from the rows under First Name and
used in the email etc.




The problem here is that I provided my colleage with an excel format of the Mail Merege column, instead in the sheet he gave me, there was no First Name  or Last Name 
column instead they were merged together under a column "Full Name" 

So I resulted to Python, to parse the Excel sheet he sent for a column "Full Name", then identify the number of strings if there are two strings, the first strings/name,
should be placed under a column called "First Name" in a new excel file, while the 2nd string/name should be placed under "Last Name" in the new excel file as well.

If there are 3 strings/name, the 1st and 2nd String/Name should be placed under "First Name" and the 3rd String/Name  will be placed under "Last Name" in the new file


NOTE !!

1.  I was fraught with some errors and I was require to install the python library openpyxl


SOFTWARE VERSION USED 

2 PIP Version 23.1.1    Python Version 3.10.11   

0S Windows 11 22H2 (Build 22621.5555)
