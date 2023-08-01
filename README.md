# new purchase order, pseudo script
### This script checks for email messages that contain keywords with information about new purchase orders.
### When the message contains all keywords, script extracts information by using regex and saves data to a JSON file.
### Finally, it connects to SAP and enters data from JSON file to create a new purchase order in SAP.
