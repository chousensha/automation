# automation
Various automation tasks

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

outlook.py

Send mails using the local Outlook profile

-c HTML formatted mail

-a attachment

Tested on Windows 10 with Outlook 2013

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

winharvest.py

Writes Excel file with the following info:
- computer name
- user name
- IP
- list of installed software and its version
- total number of installed programs

Tested on Windows XP, 7, 10

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

line_length.py

Find and print lines from a file that are a certain length

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

eyewitness-parse.py

Takes a path of Eyewitness reports, parses the generated HTML files and creates an Excel spreadsheet with the values for the domain, title page and server headers for each entry

Tested on Windows 7 with Python 2.7

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

codex.py

Takes the input and returns it encoded into ASCII character codes, URL, HTML, hex and base64. Intended for web attacks.

Drop payload here:
alert(10)
ASCII codes payload:
[97, 108, 101, 114, 116, 40, 49, 48, 41]
###############################################################
URL encoded payload: alert%2810%29
###############################################################
HTML encoded payload:
b'alert(10)'
###############################################################
Hex encoded payload:
b'616c65727428313029'
###############################################################
Base64 payload:
b'YWxlcnQoMTAp'
