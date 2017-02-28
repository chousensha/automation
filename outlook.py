import argparse
import win32com.client as win32  # install with pip install pypiwin32
"""
win32com provides access to COM objects
"""

DESC = 'Send Outlook mail'
parser = argparse.ArgumentParser(description=DESC)
# add arg: html body
parser.add_argument(
    '-c',
    help = 'Send HTML formatted mail',
    action = 'store_true' # to act as a flag, and not require anything extra after it
)

# add arg: attachment
parser.add_argument(
    '-a',
    help = 'Add attachment',
    action = 'store_true'
)

args = parser.parse_args()
argdict = vars(args)

ATT = r'some path' # raw string for Windows
outlook = win32.Dispatch('outlook.application')  # dynamic dispatch -> this is a generic COM object (Python has
# no special knowledge of this obj besides the name that you gave it

count = 1
sendrange = count + 1
while count < sendrange: #  iterate over how many times you want to send the mail
    mail = outlook.CreateItem(0)  # in VBA, 0 is for olMailItem object
    mail.To = 'recipient1; recipient2'
    mail.Subject = 'test from python # ' + str(count)
    if argdict['c'] != False:
        mail.HTMLBody = '<h2>HTML</h2>'
    elif argdict['a'] != False:
        mail.Attachments.Add(ATT)
    else:
        mail.body = 'yo yo'

    mail.send
    count += 1






