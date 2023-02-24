# ---------------------------------------------------------------------------------------------------------------------
# o365_lib_helper.py
# Description: library to interact with Microsoft 365 using o365-pyhon https://github.com/O365/python-o365
# Created 19/2/23
# Author: William Hamilton
# Last Updated:
# Author:
# Suggestions:
#
__ToDo__ = """
- add logging
"""

#
# ---------------------------------------------------------------------------------------------------------------------


import getpass
import os
import configparser

from O365 import Account, FileSystemTokenBackend

# Setup config file
current_filename = os.path.basename(__file__).split('.')[0]
config_file = f'{current_filename}.ini'
config = configparser.ConfigParser()
config.read(config_file)

# read and assign from config file
my_tenant_id = config['o365']['tenant_id']
client_id = config['o365']['client_id']
client_secret = config['o365']['client_secret']
my_protocol = config['o365']['protocol']
credentials = (client_id, client_secret)
# scopes = ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send']
current_user = getpass.getuser()
my_folder = 'mailbot_tokens'
my_token = 'mailbot_token_' + current_user + '.txt'
token_backend = FileSystemTokenBackend(token_path=my_folder, token_filename=my_token)


def authenticate_o365_current_user():
    """
    A token is read if it exists, if it is valid then this is used for authentication, if not it
    authenticates the current user with the O365 API and returns the authenticated Account object.

    The current username is used to identify token file allowing multiple users to use this script

    Returns:
        An Account object representing the authenticated user, or None if authentication failed.
    """
    my_account = Account(credentials, token_backend=token_backend)
    try:
        if not my_account.is_authenticated:
            # account = Account(credentials, auth_flow_type='credentials', tenant_id=my_tenant_id)
            my_account.authenticate(scopes=['basic', 'message_all'])
        print('Authenticated!')
    except Exception as e:
        print(f"Failed to authenticate: {e}")
        return None
    return my_account


def send_email(to_email, subject_header, msg_body):
    """
    Sends an email to the specified email address using the authenticated O365 account.

    Args:
        to_email: A string representing the email address of the recipient.
        subject_header: A string representing the subject of the email.
        msg_body: A string representing the body of the email.
    """
    try:
        m = account.new_message()
        m.to.add(to_email)
        m.subject = subject_header
        m.body = msg_body
        m.send()
        print(f"Email sent to {to_email} with subject '{subject_header}'")
    except Exception as e:
        print(f"Failed to send email: {e}")


def search_emails(subject=None, from_address=None, attachment_name=None, attachment_extension=None, date_after=None):
    """
    Searches a mailbox for an email from a particular address and with an attachment of a particular kind.

    Args:
        subject: A string representing the message subject line
        from_address: A string representing the email address of the sender.
        attachment_name: A string representing the name of the attachment to search for.
        attachment_extension: A string representing the extension of the attachment to search for.

    Returns:
        A list of Message objects matching the search criteria, or None if no matches were found.
    """
    print(attachment_extension)
    try:
        mailbox = account.mailbox()
        inbox = mailbox.inbox_folder()

        # Build the filter for the search
        query = mailbox.q().on_attribute('subject').contains(subject)
        # query.chain('and').on_attribute('receivedDateTime').greater(datetime(date_after))
        # query = mailbox.q().on_attribute('receivedDateTime').greater(datetime(2023, 2, 17))

        print(f'The query is: {query}')

        # Search the mailbox using the filter
        messages = inbox.get_messages(query=query)
        # print(f'Messages are: {messages}')
        # print(f'Message type is: {type(messages)}')
        # print(f'Attachments: {messages.attachments}')

        # Filter the results by attachment name and extension
        results = []
        for msg in messages:
            results.append(msg)

        if not results:
            print("No matching emails found.")
            return None

        return results

    except ConnectionError as e:
        print(f"Error connecting to email server: {e}")
        return None

    except Exception as e:
        print(f"Error searching emails: {e}")
        return None


if __name__ == "__main__":

    email_to = config['testing']['email_to']
    email_subject = config['testing']['email_subject']
    email_body = config['testing']['email_body']
    attachment_type = config['testing']['attachment_type']
    email_from = config['testing']['email_from']
    attachment_name = config['testing']['attachment_name']

    account = authenticate_o365_current_user()


    def test_list_inbox(subject=None):
        if account:
            mailbox = account.mailbox()
            inbox = mailbox.inbox_folder()
            # print(inbox)
            order_by = 'receivedDateTime desc'
            # query = mailbox.new_query()
            # query = mailbox.new_query('subject').startswith('t')
            query = mailbox.new_query()
            # print(f"The query is: {query}")
            msg = inbox.get_messages(query=query, order_by=order_by)

            for x in msg:
                print(x)

            # query = mailbox.q().search('testing')
            query2 = mailbox.q().on_attribute('subject').contains(subject)
            print(query2)
            messages = mailbox.inbox_folder().get_messages(query=query2)
            print('search results 2 \n')

            # for x in messages:
            #     print(x)


    def test_email_send():

        if account:
            send_email(email_to, email_subject, email_body)


    def test_email_search():
        print("Search 1 ")
        results = search_emails(subject='Check', date_after="2023, 2, 15")
        print(results)
        # print(f'type: {print(type(results[0]))}')
        # print(results[0].get_mime)
        # print(results[0]._Message__attachment)

        # messages = search_emails(subject='Check')
        # print(messages)
        # for msg in messages:
        #     print(msg.body_preview)
        # print("Search 2 ")
        # search_emails(from_address=email_from)
        # print("Search 3 ")
        # search_emails()


    # test_email_send()
    # test_email_search()
    test_list_inbox()
    # print(__ToDo__)
