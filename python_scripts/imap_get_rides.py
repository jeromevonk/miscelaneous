from backports import ssl
from imapclient import IMAPClient
import pyzmail
import imaplib
import bs4
import getpass
import calendar
from re import sub
from decimal import Decimal

imaplib._MAXLINE = 10000000

IMAP_SERVER  = 'imap.gmail.com'
USERNAME     = 'user@gmail.com'

def getRidesCost(company, month):
    AFTER  = "2017/%d/01" % (month)
    BEFORE = "2017/%d/01" % (month+1)
    
    TO_SEARCH_UBER   = "after:%s before:%s 'trip with uber'" % (AFTER,BEFORE)
    TO_SEARCH_CABIFY = "after:%s before:%s 'journey with cabify'" % (AFTER,BEFORE)

    UBER_TAG   = "td"
    UBER_CLASS = "totalPrice topPrice tal black"

    CABIFY_TAG   = "h1"
    CABIFY_CLASS = "heading-title txt-right"

    total = 0
    UIDs  = []
    print("Searching for rides with %s in %s" % (company, calendar.month_name[month]) )

    # Search
    if company == "Uber":
        UIDs = imapObj.gmail_search(TO_SEARCH_UBER)
    else:
        UIDs = imapObj.gmail_search(TO_SEARCH_CABIFY)
    
    print("%d ride(s) found" % len(UIDs ))

    # Fetch
    rawMessages = imapObj.fetch(UIDs, ['BODY[]'])

    for id in UIDs:
        message = pyzmail.PyzMessage.factory(rawMessages[id][b'BODY[]'])
        
        if message.html_part != None:
            page = bs4.BeautifulSoup(message.html_part.get_payload().decode(message.html_part.charset), "lxml")
            
            if company == "Uber":
                value = page.find_all(UBER_TAG, class_=UBER_CLASS)
            else:
                value = page.find_all(CABIFY_TAG, class_=CABIFY_CLASS)
                
            if value == []:
                print('Could not find value')
            else:
                ride = Decimal(sub(r'[^\d.]', '', value[0].contents[0]))
                print(ride)
                total += ride
                
    print("Rides with %s: %.2f\n" % (company, total))
    return total

""" Connect with IMAP"""
# IMAP Connection - workaround (http://stackoverflow.com/a/40209809/660711)
try:
    context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    imapObj = IMAPClient(IMAP_SERVER, ssl=True, ssl_context=context)
except:
    print("Exception connecting to server '%s'" % IMAP_SERVER)
    quit()

# Login
_pswd  = getpass.getpass('Password: ')

try:
    imapObj.login(USERNAME, _pswd)
except:
    print("Exception loggin in user '%s'. Wrong password?" % USERNAME)
    quit()

# Select a folder
imapObj.select_folder('INBOX', readonly=True)

by_month = []

for month in range(1,5):
    """ Pegar todas as viagens com uber no mes"""
    total_uber = getRidesCost("Uber", month)

    """ Pegar todas as viagens com cabify no mes"""
    total_cabify = getRidesCost("Cabify", month)

    by_month.append(total_uber + total_cabify) 
    print("---> Total of rides in %s: %.2f\n" % (calendar.month_name[month], by_month[-1]) )

# Printing the result
for i in range(0, len(by_month)):
    print("%s: %.2f" % (calendar.month_name[i+1], by_month[i]))    
    
# Logout
imapObj.logout()