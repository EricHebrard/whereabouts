import json
import urllib.request
from datetime import datetime, timezone

from icalendar import Calendar


CALENDARS = [
    {
        "name": "Academics",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/d7010a2bb4f245e4bc436a82091e5b6016012157490710576232/calendar.ics",
        "color": "#4285F4",
    },
    {
        "name": "Research",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/04bc370fcec948b981d8980753dd2f1313185571384744592492/calendar.ics",
        "color": "#34A853",
    },
    {
        "name": "Teaching",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/6d5bf133fa0d498196ac619ec2cc242e15629469048917954620/calendar.ics",
        "color": "#F29900",
    },
    {
        "name": "Travel",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/176d8634938d473b9a57fefd8f15ad7e15652466246061674388/calendar.ics",
        "color": "#8E44AD",
    },
]


def download_calendar(url):
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0"
        }
    )

    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read()


def convert_datetime(value):
    """
    Convert an iCalendar datetime/date into an ISO 8601 string
    suitable for FullCalendar.
    """

    if hasattr(value, "dt"):
        value = value.dt

    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.isoformat()
        return value.isoformat()

    # All-day event
    return value.isoformat()


def main():

    all_events = []

    for calendar_info in CALENDARS:

        print(f"Downloading {calendar_info['name']}...")

        data = download_calendar(calendar_info["url"])
        calendar = Calendar.from_ical(data)

        count = 0

        for component in calendar.walk():

            if component.name != "VEVENT":
                continue

            uid = str(component.get("UID", ""))

            summary = str(
                component.get("SUMMARY", "Untitled event")
            )

            dtstart = component.get("DTSTART")

            if not dtstart:
                continue

            start = convert_datetime(dtstart)

            location = str(component.get("LOCATION", ""))

            event = {
                "id": f"{calendar_info['name']}-{uid}",
                "title": summary,
                "start": start,
                "backgroundColor": calendar_info["color"],
                "borderColor": calendar_info["color"],
                "extendedProps": {
                    "calendar": calendar_info["name"]
                }
            }

            dtend = component.get("DTEND")

            if dtend:
                event["end"] = convert_datetime(dtend)

            # Preserve all-day events.
            if hasattr(dtstart.dt, "year") and not isinstance(
                dtstart.dt, datetime
            ):
                event["allDay"] = True

            all_events.append(event)
            count += 1

        print(f"  {count} events")

    with open("calendar.json", "w", encoding="utf-8") as file:
        json.dump(
            all_events,
            file,
            ensure_ascii=False,
            indent=2
        )

    print(f"Written {len(all_events)} events to calendar.json")


if __name__ == "__main__":
    main()
