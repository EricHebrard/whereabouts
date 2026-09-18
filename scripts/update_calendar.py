import json
import urllib.request
from datetime import datetime, timezone

from icalendar import Calendar


CALENDARS = [
    {
        "name": "Academics",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/153c302eae1241669e5ac7cd0678fe7c5193535583262517238/calendar.ics",
        "color": "#4285F4",
    },
    {
        "name": "Research",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/ae60a8e9f5e947eeafc1d5afed2647ab4825117944977208683/calendar.ics",
        "color": "#34A853",
    },
    {
        "name": "Teaching",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/9e1e7ceacdeb40a091a09257c68584ef9816993180113323493/calendar.ics",
        "color": "#F29900",
    },
    {
        "name": "Travel",
        "url": "https://outlook.office365.com/owa/calendar/17fdb5c3ca62463ca30184560b629747@exeter.ac.uk/b2800759e43c4aae84d6f3c4cc34982d12859872542290148978/calendar.ics",
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
                    "calendar": calendar_info["name"],
                    "location": location
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
