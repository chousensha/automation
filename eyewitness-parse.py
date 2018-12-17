from requests_testadapter import Resp
import requests
import os
from bs4 import BeautifulSoup
from collections import OrderedDict
import glob
import openpyxl


# class adapted from this Stack Overflow post:
# https://stackoverflow.com/questions/10123929/python-requests-fetch-a-file-from-a-local-url/22989322
class LocalFileAdapter(requests.adapters.HTTPAdapter):
    """
    Defining a custom adapter for opening file:// paths
    HTTPAdapter = the built-in HTTP Adapter for urllib3
    """
    def build_response_from_file(self, request):
        # get path after file://
        file_path = request.url[7:]
        with open(file_path, 'rb') as file:
            buff = bytearray(os.path.getsize(file_path))
            file.readinto(buff)
            # urllib3 response object
            resp = Resp(buff)
            # build Response object,this should not be called from user code, and is only exposed for use
            # when subclassing the HTTPAdapter
            r = self.build_response(request, resp)

            return r

    def send(self, request, stream=False, timeout=None,
             verify=True, cert=None, proxies=None):

        return self.build_response_from_file(request)

requests_session = requests.session()
requests_session.mount('file://', LocalFileAdapter())

base_path = 'C:/eyewitness/reports/path'  # CHANGE THIS TO YOUR PATH
report_files = base_path + '*'
# list the path of all report pages: report.html, report_page2.html, etc.
pages = glob.glob(report_files)
# number of report pages
page_total = len(pages)
reports = []
# add all the report pages to be handled with file://
reports.append('file://' + base_path + '.html')
# start at 2 because the html for the first page doesn't have a number
for i in range(2,page_total + 1):
    reports.append('file://' + base_path + str('_page') + str(i) + '.html')

# tracker for total entries, for debugging
count = 0
extract = OrderedDict({})

# parse HTML
for report in reports:
    source = requests_session.get(report)
    soup = BeautifulSoup(source.content, 'html.parser')
    # print(soup.prettify())


    # find the right div type that has the content
    divs = soup.find_all('div', {"style": "display: inline-block; width: 300px; word-wrap: break-word"})
    for divobj in divs:
        #print "Div #" + str(count)
        #print divobj

        # works for extracting the link
        for a in divobj.findAll('a', href=True):
            if a['href'].startswith('source'):
                pass
            else:
                domain = a['href']
                extract[domain] = []

        # categories to be extracted are in bold
        for b in divobj.findAll('b'):
            if "Resolved to" in b.text:
                #print b.nextSibling
                extract[domain].append(b.nextSibling)
            if "Page Title" in b.text:
                #print b.nextSibling.strip() # gets rid of extra newline
                extract[domain].append(b.nextSibling.strip())
            if "server" in b.text:
                #print b.nextSibling
                extract[domain].append(b.nextSibling.strip())
        count += 1


print(count)

FILE = 'report.xlsx'
def write_xls():
    '''
    Create an Excel file with the desired values
    :return:
    '''
    wbook = openpyxl.Workbook()
    sheet = wbook.get_sheet_by_name('Sheet')
    sheet['A1'] = 'DOMAIN'
    sheet.column_dimensions['A'].width = 25
    sheet['B1'] = 'IP'
    sheet['C1'] = 'PAGE TITLE'
    sheet.column_dimensions['B'].width = 25
    sheet['D1'] = 'SERVER'
    sheet.column_dimensions['C'].width = 20
    wbook.save(FILE)
    return
write_xls()

def write_to_excel(spreadsheet):
    '''
    Write values
    :param spreadsheet: Excel file to open and make changes to
    :return:
    '''
    wbook = openpyxl.load_workbook(spreadsheet)
    sheet = wbook.get_sheet_by_name('Sheet')
    start_count = 2
    for key, value in extract.iteritems():
        sheet['A' + str(start_count)] = key
        try:
            sheet['B' + str(start_count)] = value[0]
            sheet['C' + str(start_count)] = value[1]
            sheet['D' + str(start_count)] = value[2]
        except IndexError, e:
            pass
        start_count += 1
    wbook.save(spreadsheet)
    return

write_to_excel(FILE)

