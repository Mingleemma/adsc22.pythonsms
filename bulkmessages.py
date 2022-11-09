##Import Modules for networking and for Airmore messaging service and an added module from openpyxl.
## You can use any module that is able to read and work with python.

from ipaddress import IPv4Address
from pyairmore.request import AirmoreSession
from pyairmore.services.messaging import MessagingService  # to send messages
from openpyxl import load_workbook

ip = IPv4Address("192.168.137.203")  # let's create an IP address object
# now create a session
session = AirmoreSession(ip)
# if your port is not 2333
# session = AirmoreSession(ip, 2334)  # assuming it is 2334

was_accepted = session.request_authorization()

print("Is request accepted? ", was_accepted)  # True if accepted

# path to Excel Sheet
filepath = "test.xlsx"

# column to Read from
column = "A"  # suppose it is under "A"

########################
# Needs to be specified#
########################
length = 6

workbook = load_workbook(filename=filepath, read_only=True)
worksheet = workbook.active  # we will get the active worksheet

## Adding phone numbers in sheet to an array of phone numbers

phone_numbers = []
for i in range(length):
    cell = "{}{}".format(column, i + 1)
    number = worksheet[cell].value
    if number != "" or number is not None:
        phone_numbers.append(str(number))

# print (phone_numbers)

message = "Welcome to ADSC 2022."
for number in phone_numbers:
    service = MessagingService(session)
    service.send_message(number, message)
    print("message sent to " + number)