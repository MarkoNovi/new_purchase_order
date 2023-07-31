## this is just the pseudo code, dont try to use it in production



from traceback import print_exc
from exchangelib import DELEGATE, Account, Credentials, Configuration, UTC_NOW
from exchangelib import OAuth2Credentials, Version, Build, OAUTH2, IMPERSONATION, Identity
from exchangelib import EWSDateTime, EWSTimeZone
from exchangelib import Q
from datetime import datetime
import re
import json




def extract_order_details(email_content):
    '''
    Extracts order details from the given email content using regex.

    Args:
        email_content (str): The content of the email from which order details will be extracted.

    Returns:
        dict or False: A dictionary containing the extracted order details if found,
            or False if no order details match the pattern in the email content.

    Raises:
        Exception: Any exception raised during the extraction process will be caught and printed.
            The actual exception message will be displayed as the error message.
    '''

    try:
        order_details = {}
        pattern = r"Sold-to Party: (.*?)\nShip-to Party: (.*?)\nPurchase Order Number: (.*?)\nPurchase Order Date: (.*?)\nMaterial Number: (.*?)\nDescription: (.*?)\nOrdered Quantity: (.*?)\nDate of Delivery: (.*?)\n"
        match = re.search(pattern, email_content)

        if match:
            order_details['Sold-to Party'] = match.group(1)
            order_details['Ship-to Party'] = match.group(2)
            order_details['Purchase Order Number'] = match.group(3)
            order_details['Purchase Order Date'] = match.group(4)
            order_details['Material Number'] = match.group(5)
            order_details['Description'] = match.group(6)
            order_details['Ordered Quantity'] = match.group(7)
            order_details['Date of Delivery'] = match.group(8)

            return order_details
        else:
            return False
        
    except Exception as e:
        print(str(e))
        return False




def save_order_details(email_content):
    '''
    Extracts and saves order details from the email content to a JSON file.

    Args:
        email_content (str): The content of the email from which order details will be extracted.

    Raises:
        ValueError: If the order_details dictionary is empty or None, indicating that no order details were found in the email content.
    '''
    timestamp = datetime.strftime(datetime.now(), "%Y%m%d%H%M%S")
    order_details = extract_order_details(email_content)
    if order_details:
        print("Order Details:")
        for key, value in order_details.items():
            print(f"{key}: {value}")
        buyer = order_details["Sold-to Party"]

        # Save order_details to a JSON file
        with open(f'order_details_{buyer}_{timestamp}.json', 'w') as json_file:
            json.dump(order_details, json_file, indent=4)
        print("Order details saved to '.json' file.")
    else:
        print("No order details found in the email.")




def read_config_from_json(file_path):

    with open(file_path, 'r') as json_file:
        config_data = json.load(json_file)

    clientId = config_data['clientId']
    secretValue = config_data['secretValue']
    tenantId = config_data['tenantId']
    primarysa = config_data['primarysa']
    password = config_data['password']
    login_mail = config_data['login_mail'] 

    return clientId, secretValue, tenantId, primarysa, password, login_mail



def mailbox(login_mail, password, clientId, secretValue, tenantId, primarysa):
    '''
    Retrieves email messages from a mailbox folder that match the specified criteria.

    Args:
        login_mail (str): The login email address.
        password (str): The password associated with the login email address. [Note: Password is provided as an argument in the function signature but is not used in the function.]
        clientId (str): The client ID required for OAuth2 authentication.
        secretValue (str): The secret value required for OAuth2 authentication.
        tenantId (str): The tenant ID required for OAuth2 authentication.
        primarysa (str): The primary SMTP address of the mailbox.

    Returns:
        list or False: A list of email messages from the mailbox folder that match the specified criteria, or False if an error occurred

    Raises:
        Any exceptions raised during the retrieval process are not caught in this function and will be propagated to the caller
    '''
    try:

        credentials = OAuth2Credentials(client_id=clientId, client_secret=secretValue, tenant_id=tenantId, identity=Identity(primary_smtp_address=login_mail))

        ## configuration object, needs smtp address and credentials
        config = Configuration(server ='smtp.office365.com', credentials=credentials)

        ## credentials object, your login mail and password
        account = Account(
            primarysa, 
            config=config,
            access_type=DELEGATE,
        )

        inbox_folder = account.inbox
        ## get all messages in folder "inbox"
        all_messages = inbox_folder.all()
        ## get all unread messages in folder
        filter_condition = Q(is_read=False)
        filtered_messages = all_messages.filter(filter_condition)

        return filtered_messages

    except:
        print_exc()
        return False
    




def enter_data_to_sap(order_details):
    '''
    Enters order details into the SAP system using the provided SAP function module.

    Args:
        order_details (dict): A dictionary containing the order details to be entered into SAP.
            The keys in the dictionary should match the parameter names expected by the SAP function module.

    Returns:
        any or False: The result returned by the SAP function module if the data entry is successful,
            or False if an error occurred during the SAP function call.

    Raises:
        Exception: Any other exception raised during the SAP function call will be caught and printed.
            The actual exception message will be displayed as the error message.
    '''
    try:
        from pyrfc import Connection

        conn = Connection(user='sap_username', passwd='sap_password', ashost='sap_host', sysnr='sap_system_number')

        # Replace 'your_function_module_name' with the actual SAP function module to create a purchase order
        result = conn.call('your_function_module_name', **order_details)

        conn.close()
        return result
    except Exception as e:
        print(str(e))
        return False




def main_app():
    '''
    This function is the main application logic that processes emails and performs actions based on the contents.

    Parameters: None
    Return: None
    '''

    # Call the function to read the variables from the JSON file
    file_path = 'credentials.json'
    clientId, secretValue, tenantId, primarysa, password, login_mail= read_config_from_json(file_path)

    ## get unread messages from inbox
    filtered_messages = mailbox(login_mail, password, clientId, secretValue, tenantId, primarysa)
    # for item in filtered_messages:
    filtered_messages.count()
    for message in filtered_messages:

        email_content = message.text_body
        order_details = extract_order_details(email_content)

        if order_details is False: continue

        if save_order_details(order_details) is False: continue

        if enter_data_to_sap(order_details) is False: continue
        print(f"order {order_details['Sold-to Party']} is entered to sap")
        

if __name__ == "__main__":
    try:
        main_app()
    except Exception as e:
        print_exc()






