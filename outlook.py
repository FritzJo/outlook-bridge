import win32com.client
import datetime
from datetime import datetime, timedelta
import getpass
import click
from caldav_client import CaldavClient
import uuid


def create_caldav_item(outlook_element):
    template = """
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Outlook Import//CalDAV Client//EN
BEGIN:VEVENT
UID:{uid}
DTSTART:{start}
DTEND:{end}
SUMMARY:{sum}
END:VEVENT
END:VCALENDAR
"""
    eventuid = str(uuid.uuid4())
    timestart = datetime.strptime(str(outlook_element.Start)[:-6], '%Y-%m-%d %H:%M:%S')
    timestart = timestart.strftime('%Y%m%dT%H%M%S')
    timeend = datetime.strptime(str(outlook_element.End)[:-6], '%Y-%m-%d %H:%M:%S')
    timeend = timeend.strftime('%Y%m%dT%H%M%S')
    eventsum = "Test Event"

    event = template.format(uid=eventuid, start=timestart, end=timeend, sum=eventsum)
    return event


def get_outlook_appointments(dayc):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    appointments = calendar.Items

    # Restrict to items in the next 30 days
    begin = datetime.now()
    end = begin + timedelta(days=dayc);
    restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" + end.strftime("%d/%m/%Y") + "'"
    restrictedItems = appointments.Restrict(restriction)

    return restrictedItems


def caldav_insert(caldav_c, vcal):
    calendars = caldav_c.get_calendars()
    if len(calendars) > 0:
        print("Found multiple calendars:")
        for index, cal in enumerate(calendars):
            print("[" + str(index) + "] " + cal.name)
        selection = int(input("Select a calender: "))
        calendar = calendars[selection]
    else:
        calendar = calendars[0]
    print("Using calendar: ", calendar)
    caldav_c.write_caldav_event(calendar, vcal)


@click.command()
@click.option('--proxy', default="", help='URL of the http proxy')
def sync(proxy):
    # Caldav url
    Cuser = ""
    Cpassword = getpass.getpass()
    Cproxy = ""
    Curl = ""
    caldav_c = CaldavClient(Curl, Cuser, Cpassword)
    caldav_c.set_proxy(Cproxy)
    caldav_c.connect()

    events = get_outlook_appointments(1)
    print(events[0].Subject)
    vcal = create_caldav_item(events[0])
    caldav_insert(caldav_c, vcal)


if __name__ == "__main__":
    sync()
