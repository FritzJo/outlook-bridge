import win32com.client, datetime
import datetime
import caldav
from caldav.elements import dav, cdav
import getpass

def get_outlook_appointments():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    appointments = calendar.Items

    # Restrict to items in the next 30 days
    begin = datetime.datetime.now()
    end = begin + datetime.timedelta(days = 1);
    restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" +end.strftime("%d/%m/%Y") + "'"
    restrictedItems = appointments.Restrict(restriction)

    # Iterate through restricted AppointmentItems and print them
    for appointmentItem in restrictedItems:
        print("{0} Start: {1}, End: {2}, Organizer: {3}".format(
              appointmentItem.Subject, appointmentItem.Start, 
              appointmentItem.End, appointmentItem.Organizer))
                  
def caldav_insert():
    # Caldav url
    Cuser = "
    Cpassword = getpass.getpass()
    Cproxy = ""
    Curl = ""

    vcal = """BEGIN:VCALENDAR
    VERSION:2.0
    PRODID:-//Example Corp.//CalDAV Client//EN
    BEGIN:VEVENT
    UID:1234567890
    DTSTAMP:20100510T182145Z
    DTSTART:20100512T170000Z
    DTEND:20100512T180000Z
    SUMMARY:This is an event
    END:VEVENT
    END:VCALENDAR
    """

    client = caldav.DAVClient(proxy= Cproxy, url=Curl, username=Cuser, password=Cpassword)
    principal = client.principal()
    calendars = principal.calendars()
    if len(calendars) > 0:
        print("Found multiple calendars:")
        for index, cal in enumerate(calendars):
            print("[" + str(index) + "] " + cal.name)
        selection = int(input("Select a calender: "))
        calendar = calendars[selection]
    else:
        calendar = calendars[0]
    print("Using calendar: ", calendar)
    #print "Renaming"
    #calendar.set_properties([dav.DisplayName("Test calendar"),])
    #print calendar.get_properties([dav.DisplayName(),])

    #event = calendar.add_event(vcal)
    #print "Event", event, "created"

    print("Looking for events in 2019-01")
    results = calendar.date_search(
        datetime(2020, 1, 1), datetime(2020, 1, 30))

    for event in results:
        print("Found", event)

get_outlook_appointments()
caldav_insert()
