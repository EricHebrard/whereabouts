import json
import urllib.request
from datetime import datetime, date

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

    with urllib.request.urlopen(
        request,
        timeout=30
    ) as response:
        return response.read()


def convert_value(value):
    """
    Convert an iCalendar property value into something
    that can safely be written to JSON.
    """

    if value is None:
        return None

    if hasattr(value, "dt"):
        value = value.dt

    if isinstance(value, datetime):
        return value.isoformat()

    if isinstance(value, date):
        return value.isoformat()

    if isinstance(value, bytes):
        return value.decode(
            "utf-8",
            errors="replace"
        )

    if isinstance(value, (list, tuple)):
        return [
            convert_value(item)
            for item in value
        ]

    if isinstance(value, dict):
        return {
            str(key): convert_value(item)
            for key, item in value.items()
        }

    return str(value)


def convert_datetime(value):
    """
    Convert an iCalendar datetime/date into an ISO 8601
    string suitable for FullCalendar.
    """

    return convert_value(value)


def get_property(component, name, default=None):
    """
    Safely retrieve an iCalendar property.
    """

    value = component.get(name)

    if value is None:
        return default

    return convert_value(value)


def get_attendees(component):
    """
    Extract attendee information in a useful JSON format.
    """

    attendees = []

    raw_attendees = component.get("ATTENDEE")

    if not raw_attendees:
        return attendees

    if not isinstance(raw_attendees, list):
        raw_attendees = [raw_attendees]

    for attendee in raw_attendees:

        email = str(attendee)

        if email.lower().startswith("mailto:"):
            email = email[7:]

        params = getattr(attendee, "params", {})

        attendees.append({
            "email": email,
            "name": str(params.get("CN", "")),
            "role": str(params.get("ROLE", "")),
            "status": str(params.get("PARTSTAT", "")),
            "rsvp": str(params.get("RSVP", "")),
        })

    return attendees


def get_organizer(component):
    """
    Extract organiser information.
    """

    organizer = component.get("ORGANIZER")

    if not organizer:
        return ""

    value = str(organizer)

    if value.lower().startswith("mailto:"):
        value = value[7:]

    params = getattr(organizer, "params", {})

    name = str(params.get("CN", ""))

    if name:
        return f"{name} <{value}>"

    return value


def get_categories(component):
    """
    Extract CATEGORIES as a list.
    """

    categories = component.get("CATEGORIES")

    if not categories:
        return []

    value = categories.to_ical().decode(
        "utf-8",
        errors="replace"
    )

    return [
        item.strip()
        for item in value.split(",")
        if item.strip()
    ]


def get_extra_properties(component):
    """
    Preserve additional VEVENT properties that are not
    explicitly mapped into the FullCalendar event.

    This means information isn't silently discarded merely
    because we don't currently have a dedicated UI field for it.
    """

    known = {
        "UID",
        "SUMMARY",
        "DTSTART",
        "DTEND",
        "LOCATION",
        "DESCRIPTION",
        "URL",
        "STATUS",
        "TRANSP",
        "CATEGORIES",
        "PRIORITY",
        "ORGANIZER",
        "ATTENDEE",
    }

    extra = {}

    for key, value in component.property_items():

        key = str(key)

        if key in known:
            continue

        converted = convert_value(value)

        if key in extra:

            if not isinstance(extra[key], list):
                extra[key] = [extra[key]]

            extra[key].append(converted)

        else:
            extra[key] = converted

    return extra


def main():

    all_events = []

    for calendar_info in CALENDARS:

        print(
            f"Downloading {calendar_info['name']}..."
        )

        data = download_calendar(
            calendar_info["url"]
        )

        calendar = Calendar.from_ical(data)

        count = 0

        for component in calendar.walk():

            if component.name != "VEVENT":
                continue

            uid = str(
                component.get("UID", "")
            )

            summary = str(
                component.get(
                    "SUMMARY",
                    "Untitled event"
                )
            )

            dtstart = component.get("DTSTART")

            if not dtstart:
                continue

            start = convert_datetime(dtstart)

            location = get_property(
                component,
                "LOCATION",
                ""
            )

            description = get_property(
                component,
                "DESCRIPTION",
                ""
            )

            url = get_property(
                component,
                "URL",
                ""
            )

            event = {
                "id": (
                    f"{calendar_info['name']}-{uid}"
                ),

                "title": summary,

                "start": start,

                "backgroundColor":
                    calendar_info["color"],

                "borderColor":
                    calendar_info["color"],

                "extendedProps": {

                    "calendar":
                        calendar_info["name"],

                    "location":
                        location,

                    "description":
                        description,

                    "url":
                        url,

                    "status":
                        get_property(
                            component,
                            "STATUS",
                            ""
                        ),

                    "transparency":
                        get_property(
                            component,
                            "TRANSP",
                            ""
                        ),

                    "categories":
                        get_categories(component),

                    "priority":
                        get_property(
                            component,
                            "PRIORITY",
                            ""
                        ),

                    "organizer":
                        get_organizer(component),

                    "attendees":
                        get_attendees(component),

                    "extra":
                        get_extra_properties(component),
                }
            }

            dtend = component.get("DTEND")

            if dtend:
                event["end"] = convert_datetime(dtend)

            /*
             * Preserve all-day events.
             */
            if (
                hasattr(dtstart.dt, "year")
                and not isinstance(
                    dtstart.dt,
                    datetime
                )
            ):
                event["allDay"] = True

            all_events.append(event)
            count += 1

        print(
            f"  {count} events"
        )

    with open(
        "calendar.json",
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            all_events,
            file,
            ensure_ascii=False,
            indent=2
        )

    print(
        f"Written {len(all_events)} events to calendar.json"
    )


if __name__ == "__main__":
    main()
