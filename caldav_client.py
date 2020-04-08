import caldav


class Caldav_client:
    def __init__(self, url, user, password):
        self.caldav_url = url
        self.username = user
        self.password = password

    def set_proxy(self, proxy_url):
        self.proxy_url = proxy_url

    def use_proxy(self):
        return hasattr(self, 'proxy_url')

    def connect(self):
        if self.use_proxy:
            client = caldav.DAVClient(proxy=self.proxy_url, url=self.caldav_url, username=self.username,
                                      password=self.password)
        else:
            client = caldav.DAVClient(url=self.caldav_url, username=self.username, password=self.password)
        self.principal = client.principal()

    def get_calendars(self):
        return self.principal.calendars()

    def write_caldav_event(self, calendar, vcal):
        event = calendar.add_event(vcal)
        print("Event", event, "created")
