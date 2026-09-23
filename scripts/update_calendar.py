#!/usr/bin/env python3

import json
import urllib.request
from datetime import date, timedelta, datetime

import icalendar
import recurring_ical_events


# ============================================================
# OUTLOOK CALENDAR URLS
# ============================================================
#
# IMPORTANT:
# Replace these four placeholders with the FOUR URLs from
# your existing working update_calendar.py.
#
CALENDARS = {
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
}


OUTPUT_FILE = "calendar.json"


# ============================================================
# RECURRENCE EXPANSION RANGE
# ============================================================
#
# Outlook stores a recurring meeting as a single VEVENT
# containing an RRULE.
#
# We expand that VEVENT into individual occurrences over
# this rolling period.
#
# One year backwards
# Two years forwards
#
EXPANSION_START = (
    date.today() - timedelta(days=365)
)

EXPANSION_END = (
    date.today() + timedelta(days=730)
)


# ============================================================
# DOWNLOAD CALENDAR
# ============================================================

def download_calendar(url):

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent":
                "Mozilla/5.0 "
                "(compatible; whereabouts-calendar/1.0)"
        }
    )

    with urllib.request.urlopen(
        request,
        timeout=60
    ) as response:

        return response.read()


# ============================================================
# CONVERT ICALENDAR VALUES TO JSON-FRIENDLY TEXT
# ============================================================

def value_to_string(value):

    if value is None:
        return None


    if isinstance(value, (datetime, date)):

        return value.isoformat()


    if hasattr(value, "dt"):

        try:

            dt = value.dt

            if isinstance(
                dt,
                (datetime, date)
            ):

                return dt.isoformat()

        except Exception:
            pass


    if hasattr(value, "to_ical"):

        try:

            raw = value.to_ical()

            if isinstance(raw, bytes):

                return raw.decode(
                    "utf-8",
                    errors="replace"
                )

            return str(raw)

        except Exception:
            pass


    return str(value)


# ============================================================
# GET A SINGLE ICALENDAR PROPERTY
# ============================================================

def get_property(
    component,
    name
):

    value = component.get(name)

    if isinstance(value, list):

        return [
            value_to_string(item)
            for item in value
        ]

    return value_to_string(value)


# ============================================================
# CATEGORIES
# ============================================================

def get_categories(component):

    value = component.get("CATEGORIES")

    if value is None:
        return []


    try:

        return [
            str(item)
            for item in value.cats
        ]

    except Exception:
        pass


    if isinstance(value, list):

        return [
            value_to_string(item)
            for item in value
        ]


    return [
        item.strip()
        for item in value_to_string(
            value
        ).split(",")
        if item.strip()
    ]


# ============================================================
# ORGANIZER
# ============================================================

def get_organizer(component):

    organizer = component.get(
        "ORGANIZER"
    )

    if organizer is None:
        return None


    text = value_to_string(
        organizer
    )


    common_name = None

    if hasattr(
        organizer,
        "params"
    ):

        common_name = organizer.params.get(
            "CN"
        )


    if common_name:

        return (
            f"{common_name} <{text}>"
        )


    return text


# ============================================================
# ATTENDEES
# ============================================================

def get_attendees(component):

    attendees = component.get(
        "ATTENDEE"
    )

    if attendees is None:
        return []


    if not isinstance(
        attendees,
        list
    ):

        attendees = [
            attendees
        ]


    result = []


    for attendee in attendees:

        text = value_to_string(
            attendee
        )


        if hasattr(
            attendee,
            "params"
        ):

            params = attendee.params

        else:

            params = {}


        common_name = params.get(
            "CN"
        )

        role = params.get(
            "ROLE"
        )

        partstat = params.get(
            "PARTSTAT"
        )


        if common_name:

            label = (
                f"{common_name} <{text}>"
            )

        else:

            label = text


        if role:

            label += (
                f" [{role}]"
            )


        if partstat:

            label += (
                f" ({partstat})"
            )


        result.append(label)


    return result


# ============================================================
# PRESERVE OTHER ICALENDAR PROPERTIES
# ============================================================
#
# This deliberately preserves fields that are not explicitly
# mapped above, including:
#
#   RRULE
#   RDATE
#   EXDATE
#   RECURRENCE-ID
#   DTSTAMP
#   CREATED
#   LAST-MODIFIED
#   SEQUENCE
#   Outlook X-* properties
#
# ============================================================

def get_extra_properties(component):

    known = {
        "UID",
        "SUMMARY",
        "DTSTART",
        "DTEND",
        "DURATION",
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


    for name, value in (
        component.property_items()
    ):

        if name in known:
            continue


        converted = value_to_string(
            value
        )


        if name in extra:

            if not isinstance(
                extra[name],
                list
            ):

                extra[name] = [
                    extra[name]
                ]


            extra[name].append(
                converted
            )

        else:

            extra[name] = converted


    # --------------------------------------------------------
    # Preserve VALARM components too.
    # --------------------------------------------------------

    alarms = []


    for subcomponent in getattr(
        component,
        "subcomponents",
        []
    ):

        if subcomponent.name != "VALARM":
            continue


        alarm_data = {}


        for name, value in (
            subcomponent.property_items()
        ):

            alarm_data[name] = (
                value_to_string(value)
            )


        alarms.append(
            alarm_data
        )


    if alarms:

        extra["VALARM"] = alarms


    return extra


# ============================================================
# CONVERT A VEVENT INTO A FULLCALENDAR EVENT
# ============================================================

def component_to_event(
    component,
    calendar_name
):

    dtstart_property = (
        component.get("DTSTART")
    )


    if dtstart_property is None:
        return None


    start = dtstart_property.dt


    # A date rather than datetime means
    # this is an all-day event.
    all_day = not isinstance(
        start,
        datetime
    )


    # --------------------------------------------------------
    # End time
    # --------------------------------------------------------

    dtend_property = (
        component.get("DTEND")
    )


    if dtend_property is not None:

        end = dtend_property.dt


    else:

        duration_property = (
            component.get("DURATION")
        )


        if duration_property is not None:

            end = (
                start +
                duration_property.dt
            )

        else:

            end = None


    # --------------------------------------------------------
    # UNIQUE EVENT ID
    # --------------------------------------------------------
    #
    # A recurring Outlook event normally has one UID for the
    # entire series. FullCalendar needs separate IDs for the
    # individual occurrences.
    #
    # RECURRENCE-ID identifies an overridden occurrence.
    # DTSTART provides the fallback.
    #
    # --------------------------------------------------------

    uid = (
        get_property(
            component,
            "UID"
        )
        or "no-uid"
    )


    recurrence_id = (
        get_property(
            component,
            "RECURRENCE-ID"
        )
    )


    occurrence_key = (
        recurrence_id
        or value_to_string(start)
    )


    event_id = (
        f"{calendar_name}:"
        f"{uid}:"
        f"{occurrence_key}"
    )


    # --------------------------------------------------------
    # EVENT
    # --------------------------------------------------------

    event = {

        "id": event_id,

        "title":
            get_property(
                component,
                "SUMMARY"
            )
            or "(No title)",

        "start":
            value_to_string(start),

        "allDay":
            all_day,


        "extendedProps": {

            "calendar":
                calendar_name,

            "location":
                get_property(
                    component,
                    "LOCATION"
                ),

            "description":
                get_property(
                    component,
                    "DESCRIPTION"
                ),

            "url":
                get_property(
                    component,
                    "URL"
                ),

            "status":
                get_property(
                    component,
                    "STATUS"
                ),

            "transp":
                get_property(
                    component,
                    "TRANSP"
                ),

            "categories":
                get_categories(
                    component
                ),

            "priority":
                get_property(
                    component,
                    "PRIORITY"
                ),

            "organizer":
                get_organizer(
                    component
                ),

            "attendees":
                get_attendees(
                    component
                ),

            "extra":
                get_extra_properties(
                    component
                ),
        },
    }


    if end is not None:

        event["end"] = (
            value_to_string(end)
        )


    return event


# ============================================================
# EXPAND RECURRING EVENTS
# ============================================================
#
# This is the main fix.
#
# recurring_ical_events expands:
#
#   RRULE
#   RDATE
#   EXDATE
#   recurrence exceptions
#
# into individual VEVENT occurrences.
#
# ============================================================

def expand_events(
    calendar,
    calendar_name
):

    query = (
        recurring_ical_events.of(
            calendar,
            keep_recurrence_attributes=True
        )
    )


    occurrences = query.between(
        EXPANSION_START,
        EXPANSION_END
    )


    events = []


    for component in occurrences:

        event = component_to_event(
            component,
            calendar_name
        )


        if event is not None:

            events.append(event)


    return events


# ============================================================
# PROCESS ALL FOUR CALENDARS
# ============================================================

all_events = []


for calendar_name, url in (
    CALENDARS.items()
):

    print(
        f"Downloading "
        f"{calendar_name} calendar..."
    )


    if (
        not url
        or "PASTE_" in url
    ):

        raise RuntimeError(
            f"The ICS URL for "
            f"{calendar_name} has not "
            f"been configured. "
            f"Copy the existing URL "
            f"from your current script."
        )


    data = download_calendar(
        url
    )


    calendar = (
        icalendar.Calendar.from_ical(
            data
        )
    )


    all_events.extend(
        expand_events(
            calendar,
            calendar_name
        )
    )


# ============================================================
# STABLE SORTING
# ============================================================

all_events.sort(
    key=lambda event: (
        event.get(
            "start",
            ""
        ),

        event.get(
            "end",
            ""
        ),

        event.get(
            "title",
            ""
        ),

        event.get(
            "id",
            ""
        ),
    )
)


# ============================================================
# WRITE JSON
# ============================================================

with open(
    OUTPUT_FILE,
    "w",
    encoding="utf-8"
) as output:

    json.dump(
        all_events,
        output,
        ensure_ascii=False,
        indent=2
    )


print(
    f"Wrote "
    f"{len(all_events)} event occurrences "
    f"to {OUTPUT_FILE} "
    f"for {EXPANSION_START} "
    f"through {EXPANSION_END}."
)
