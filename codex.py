import urllib.parse
import html
import binascii
import base64


def strToAscii(payload):
    encoded = []
    for char in payload:
        encoded.append(ord(char))
    return encoded


def urlEncode(payload):
    return urllib.parse.quote(payload)

def htmlEncode(payload):
    #return html.escape(payload)
    encoded = html.escape(payload)
    return encoded.encode('ascii', 'xmlcharrefreplace')


def hexEncode(payload):
    hex_str = binascii.hexlify(payload.encode('ascii'))
    return hex_str

def base64Encode(payload):
    return base64.b64encode(payload.encode('ascii'))


encode = input('Drop payload here:\n')


if __name__ == "__main__":
    print('ASCII codes payload:')
    print(strToAscii(encode))
    print('###############################################################')
    print('URL encoded payload: ' + urlEncode(encode))
    print('###############################################################')
    print('HTML encoded payload:')
    print(htmlEncode(encode))
    print('###############################################################')
    print('Hex encoded payload:')
    print(hexEncode(encode))
    print('###############################################################')
    print('Base64 payload:')
    print(base64Encode(encode))
