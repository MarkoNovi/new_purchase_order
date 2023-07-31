# new purchase order


### This script checks for email messages that containd keywords with information about new purchase order.
When message contains all keywords, then it extracts information by using regex and saves data to JSON file.
Finaly it connects to SAP and enters data form JSON file to create new purchase order in SAP.
