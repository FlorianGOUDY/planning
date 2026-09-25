#!/usr/bin/env python3
"""
Planning ITECH 1 (Excel) + fichier des salles (Excel) -> calendriers .ics

Calendars generated:

    G1
    G1-A
    G1-B

    G2
    G2-A
    G2-B

    G3
    G3-A
    G3-B

Usage:
    python planning_ics.py --download --tous --outdir docs

    python planning_ics.py PLANNING.xlsx SALLES.xlsx --tous

    python planning_ics.py PLANNING.xlsx SALLES.xlsx --groupe G2

    python planning_ics.py PLANNING.xlsx SALLES.xlsx --groupe G2-A

    python planning_ics.py PLANNING.xlsx SALLES.xlsx --groupe G2-B

Dependency:
    pip install openpyxl
"""

import argparse
import datetime as dt
import hashlib
import os
import re
import unicodedata
import urllib.request

from collections import defaultdict
from openpyxl import load_workbook


# -----------------------------------------------------------------
# SETTINGS
# -----------------------------------------------------------------

FIRST_MONDAY = dt.date(2026, 8, 24)

FIRST_COL = 3
COLS_PER_WEEK = 9

FIRST_ROW = 5
ROWS_PER_DAY = 21

DAY_START_MIN = 8 * 60

# Each group has exactly 3 columns.
#
# IMPORTANT:
#
#   column 0 -> A
#   column 1 -> A
#   column 2 -> B
#
# So the layout is always:
#
#   A | A | B
#
GROUP_COLS = {
    "G1": (0, 2),
    "G2": (3, 5),
    "G3": (6, 8),
}

ROOM_COLS = range(3, 28)

TZ = "Europe/Paris"

PLANNING_ID = (
    "1ZfgKO-68K1ioRlTjY2x6Z5C8sH1TpJ19nxcdbMMIyPI"
)

SALLES_ID = (
    "1es1RFXRtJGeyL8MzFaw6fyhEzXDbumx6GTGYiUxBSAg"
)

MIN_EVENTS = 50


# -----------------------------------------------------------------
# TEXT NORMALIZATION
# -----------------------------------------------------------------

STOP = {
    "td",
    "tp",
    "ds",
    "cm",
    "de",
    "des",
    "du",
    "la",
    "le",
    "les",
    "l",
    "d",
    "a",
    "au",
    "et",
    "en",
    "itech",
    "1",
    "2",
    "3",
    "e",
    "learning",
    "initiation",
    "introduction",
}

SYN = {
    "legislation": "droit",
    "travail": "droit",
    "economique": "eco",
    "economie": "eco",
    "resistance": "rdm",
    "materiaux": "rdm",
    "fluides": "flu",
    "mecanique": "meca",
    "thermodynamique": "thermo",
    "organique": "orga",
    "incertitudes": "incertitude",
    "analyses": "analyse",
    "instrumentales": "instrumentale",
    "scientifiques": "bsi",
    "bases": "bsi",
    "ingenieur": "bsi",
    "outils": "oin",
    "numeriques": "oin",
    "informatiques": "oin",
    "generale": "chimie",
    "profil": "profil",
    "parole": "parole",
}


def norm(s):
    s = (
        unicodedata
        .normalize("NFKD", str(s))
        .encode("ascii", "ignore")
        .decode()
        .lower()
    )

    s = re.sub(r"\s*&\s*", "&", s)

    return re.sub(
        r"[^a-z0-9&]+",
        " ",
        s,
    ).strip()


def tokens(title):
    out = set()

    for word in norm(title).split():

        if word in STOP or word.isdigit():
            continue

        out.add(
            SYN.get(word, word)
        )

    return out


def tok_match(a, b):
    return any(
        x == y
        or (
            len(x) >= 3
            and len(y) >= 3
            and (
                x.startswith(y)
                or y.startswith(x)
            )
        )
        for x in a
        for y in b
    )


# -----------------------------------------------------------------
# TEACHERS
# -----------------------------------------------------------------

TEACHER_RE = re.compile(
    r"^(?:[A-Z]{1,2}\.\s?|[A-Z]\s)"
    r"([A-ZÀ-Ý][a-zà-ÿ'\-]{2,})"
)


def teachers(lines):
    """
    Detect teacher names such as:

        J.Champliaud
        AC.Besson
        L Scalone
    """

    out = set()

    for line in lines:

        match = TEACHER_RE.match(
            line.strip()
        )

        if match:
            out.add(
                norm(match.group(1))
            )

    return out


def content_lines(lines):
    """
    Keep course-related lines.
    """

    return [
        line
        for line in lines
        if not TEACHER_RE.match(line)
        and not re.fullmatch(
            r"[-\s\d/]+",
            line,
        )
    ]


def group_marker(text):
    match = re.search(
        r"\b(?:g|grp|groupe)\s*([123])",
        norm(text),
    )

    return (
        f"G{match.group(1)}"
        if match
        else None
    )


# -----------------------------------------------------------------
# WEEK DATES
# -----------------------------------------------------------------

def week_starts(ws, n=60):
    """
    Read Monday dates from row 2.
    """

    raw = {}

    for k in range(n):

        value = ws.cell(
            2,
            FIRST_COL + COLS_PER_WEEK * k,
        ).value

        if (
            isinstance(value, dt.datetime)
            and value.year >= 2025
            and value.weekday() == 0
        ):
            raw[k] = value.date()

    if raw:

        first_index = min(raw)

        anchor = (
            raw[first_index]
            - dt.timedelta(
                days=7 * first_index
            )
        )

    else:
        anchor = FIRST_MONDAY

    return [
        raw.get(
            k,
            anchor + dt.timedelta(days=7 * k),
        )
        for k in range(n)
    ]


# -----------------------------------------------------------------
# A / B DETECTION
# -----------------------------------------------------------------

def event_subgroups(group, p1, p2):
    """
    Determine whether an event belongs to A, B or both.

    Every group has this fixed structure:

        A | A | B

    Therefore:

        first column  -> A
        second column -> A
        third column  -> B

    The function works with normal cells AND merged cells.

    Examples:

        A column only:
            -> {"A"}

        A + A merged:
            -> {"A"}

        B column:
            -> {"B"}

        A + B:
            -> {"A", "B"}

        Full group:
            -> {"A", "B"}
    """

    if group not in GROUP_COLS:
        return set()

    group_start, group_end = GROUP_COLS[group]

    overlap_start = max(
        p1,
        group_start,
    )

    overlap_end = min(
        p2,
        group_end,
    )

    if overlap_start > overlap_end:
        return set()

    result = set()

    # Relative columns inside the group:
    #
    # 0 -> A
    # 1 -> A
    # 2 -> B
    #
    for column in range(
        overlap_start,
        overlap_end + 1,
    ):

        relative = (
            column - group_start
        )

        if relative in (0, 1):
            result.add("A")

        elif relative == 2:
            result.add("B")

    return result


# -----------------------------------------------------------------
# PLANNING READER
# -----------------------------------------------------------------

def read_planning(path):

    ws = load_workbook(
        path,
        data_only=True,
    ).worksheets[0]

    starts = week_starts(ws)

    merged = {
        (m.min_row, m.min_col): m
        for m in ws.merged_cells.ranges
    }

    events = []

    max_row = (
        FIRST_ROW
        + 5 * ROWS_PER_DAY
        - 1
    )

    for row in ws.iter_rows(
        min_row=FIRST_ROW,
        max_row=max_row,
        min_col=FIRST_COL,
    ):

        for cell in row:

            if (
                cell.value is None
                or not str(cell.value).strip()
            ):
                continue

            merge = merged.get(
                (
                    cell.row,
                    cell.column,
                )
            )

            if merge:

                r2 = merge.max_row
                c2 = merge.max_col

            else:

                r2 = cell.row
                c2 = cell.column

            # -----------------------------------------------------
            # WEEK / COLUMNS
            # -----------------------------------------------------

            week = (
                cell.column - FIRST_COL
            ) // COLS_PER_WEEK

            p1 = (
                cell.column - FIRST_COL
            ) % COLS_PER_WEEK

            p2 = (
                c2 - FIRST_COL
            ) % COLS_PER_WEEK

            if (
                c2 - cell.column
                >= COLS_PER_WEEK
            ):
                p2 = COLS_PER_WEEK - 1

            # -----------------------------------------------------
            # DAY / TIME
            # -----------------------------------------------------

            day, offset = divmod(
                cell.row - FIRST_ROW,
                ROWS_PER_DAY,
            )

            n_rows = (
                r2 - cell.row + 1
            )

            # -----------------------------------------------------
            # GROUPS
            # -----------------------------------------------------

            groups = [
                group
                for group, (a, b)
                in GROUP_COLS.items()
                if p1 <= b
                and p2 >= a
            ]

            full_groups = [
                group
                for group, (a, b)
                in GROUP_COLS.items()
                if p1 <= a
                and p2 >= b
            ]

            partial = any(
                group not in full_groups
                for group in groups
            )

            # -----------------------------------------------------
            # SUBGROUPS
            # -----------------------------------------------------

            subgroup_map = {}

            for group in groups:

                subgroup_map[group] = (
                    event_subgroups(
                        group,
                        p1,
                        p2,
                    )
                )

            # -----------------------------------------------------
            # DATE
            # -----------------------------------------------------

            base_date = (
                starts[week]
                + dt.timedelta(days=day)
            )

            text = str(
                cell.value
            ).strip()

            # -----------------------------------------------------
            # ALL-DAY EVENT
            # -----------------------------------------------------

            if (
                offset == 0
                and n_rows >= ROWS_PER_DAY
            ):

                ndays = max(
                    1,
                    min(
                        5 - day,
                        round(
                            n_rows
                            / ROWS_PER_DAY
                        ),
                    ),
                )

                events.append(
                    dict(
                        date=base_date,
                        end_date=(
                            base_date
                            + dt.timedelta(
                                days=ndays
                            )
                        ),
                        allday=True,
                        text=text,
                        groups=groups,
                        subgroups=subgroup_map,
                        partial=partial,
                    )
                )

                continue

            # -----------------------------------------------------
            # NORMAL EVENT
            # -----------------------------------------------------

            r2 = min(
                r2,
                FIRST_ROW
                + (day + 1)
                * ROWS_PER_DAY
                - 1,
            )

            start = (
                DAY_START_MIN
                + 30 * offset
            )

            end = (
                DAY_START_MIN
                + 30 * (
                    r2
                    - (
                        FIRST_ROW
                        + day
                        * ROWS_PER_DAY
                    )
                    + 1
                )
            )

            # -----------------------------------------------------
            # EXPLICIT TIME IN TEXT
            # -----------------------------------------------------

            time_match = re.search(
                r"(\d{1,2})\s*h\s*(\d{2})?"
                r"\s*-\s*"
                r"(\d{1,2})\s*h\s*(\d{2})?",
                text,
            )

            if time_match:

                start2 = (
                    int(
                        time_match.group(1)
                    )
                    * 60
                    + int(
                        time_match.group(2)
                        or 0
                    )
                )

                end2 = (
                    int(
                        time_match.group(3)
                    )
                    * 60
                    + int(
                        time_match.group(4)
                        or 0
                    )
                )

                if (
                    6 * 60
                    <= start2
                    < end2
                    <= 22 * 60
                ):
                    start = start2
                    end = end2

            events.append(
                dict(
                    date=base_date,
                    start=start,
                    end=end,
                    allday=False,
                    text=text,
                    groups=groups,
                    subgroups=subgroup_map,
                    partial=partial,
                )
            )

    return events


# -----------------------------------------------------------------
# ROOMS READER
# -----------------------------------------------------------------

def read_rooms(path):

    ws = load_workbook(
        path,
        data_only=True,
    )["Salles de cours"]

    headers = {
        column: str(
            ws.cell(
                3,
                column,
            ).value
            or ""
        ).split("\n")[0].strip()
        for column in ROOM_COLS
    }

    merged = {
        (m.min_row, m.min_col): m
        for m in ws.merged_cells.ranges
    }

    output = []

    for row in range(
        4,
        ws.max_row + 1,
    ):

        date_value = ws.cell(
            row,
            1,
        ).value

        if not isinstance(
            date_value,
            dt.datetime,
        ):
            continue

        for room_row in range(
            row,
            row + ROWS_PER_DAY,
        ):

            for column in ROOM_COLS:

                value = ws.cell(
                    room_row,
                    column,
                ).value

                if (
                    value is None
                    or not str(value).strip()
                ):
                    continue

                merge = merged.get(
                    (
                        room_row,
                        column,
                    )
                )

                room_row_end = min(
                    (
                        merge.max_row
                        if merge
                        else room_row
                    ),
                    row
                    + ROWS_PER_DAY
                    - 1,
                )

                output.append(
                    dict(
                        date=date_value.date(),
                        start=(
                            DAY_START_MIN
                            + 30
                            * (
                                room_row
                                - row
                            )
                        ),
                        end=(
                            DAY_START_MIN
                            + 30
                            * (
                                room_row_end
                                - row
                                + 1
                            )
                        ),
                        room=headers[column],
                        text=str(
                            value
                        ).strip(),
                    )
                )

    return output


# -----------------------------------------------------------------
# PLANNING <-> ROOMS MATCHING
# -----------------------------------------------------------------

def find_rooms(
    event,
    group,
    rooms_by_date,
):

    if event["allday"]:
        return []

    lines = [
        line.strip()
        for line in event["text"].split("\n")
        if line.strip()
    ]

    if not lines:
        return []

    if norm(
        lines[0]
    ).startswith(
        "e learning"
    ):
        return []

    content = content_lines(
        lines
    )

    if content:

        title_tokens = set().union(
            *[
                tokens(line)
                for line in content[:2]
            ]
        )

    else:
        title_tokens = set()

    if not title_tokens:
        return []

    event_teachers = teachers(
        lines
    )

    found = []

    for room in rooms_by_date.get(
        event["date"],
        [],
    ):

        room_lines = [
            line.strip()
            for line in room["text"].split("\n")
            if line.strip()
        ]

        if not re.search(
            r"itech\s*1",
            norm(room["text"]),
        ) and not (
            event_teachers
            & teachers(room_lines)
        ):
            continue

        overlap = (
            min(
                event["end"],
                room["end"],
            )
            - max(
                event["start"],
                room["start"],
            )
        )

        if overlap < (
            0.6
            * min(
                event["end"]
                - event["start"],
                room["end"]
                - room["start"],
            )
        ):
            continue

        room_content = content_lines(
            room_lines
        )

        room_tokens = (
            set().union(
                *[
                    tokens(line)
                    for line in room_content[:1]
                ]
            )
            if room_content
            else set()
        )

        if not tok_match(
            title_tokens,
            room_tokens,
        ):
            continue

        marker = group_marker(
            room["text"]
        )

        if (
            marker
            and group
            and marker != group
        ):
            continue

        room_teachers = teachers(
            room_lines
        )

        if (
            event_teachers
            and room_teachers
            and not (
                event_teachers
                & room_teachers
            )
        ):
            continue

        found.append(
            dict(
                room=room["room"],
                teacher=", ".join(
                    sorted(
                        teacher.title()
                        for teacher
                        in room_teachers
                    )
                )
                if room_teachers
                else "",
                marked=bool(marker),
            )
        )

    if any(
        item["marked"]
        for item in found
    ):

        found = [
            item
            for item in found
            if item["marked"]
        ]

    unique = []
    seen = set()

    for item in found:

        if item["room"] in seen:
            continue

        seen.add(
            item["room"]
        )

        unique.append(
            item
        )

    return unique


# -----------------------------------------------------------------
# ICS HELPERS
# -----------------------------------------------------------------

def esc(value):

    return (
        str(value)
        .replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\n", "\\n")
    )


def fold(line):

    data = line.encode(
        "utf-8"
    )

    output = []

    while len(data) > 74:

        cut = 74

        while (
            cut > 0
            and (
                data[cut]
                & 0xC0
            ) == 0x80
        ):
            cut -= 1

        output.append(
            data[:cut].decode(
                "utf-8"
            )
        )

        data = data[cut:]

        data = (
            b" "
            + data
        )

    output.append(
        data.decode("utf-8")
    )

    return "\r\n".join(
        output
    )


VTZ = """
BEGIN:VTIMEZONE
TZID:Europe/Paris
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE
""".strip().replace(
    "\n",
    "\r\n",
)


# -----------------------------------------------------------------
# CALENDAR SCOPE
# -----------------------------------------------------------------

def parse_calendar_scope(scope):

    match = re.fullmatch(
        r"(G[123])(?:-([AB]))?",
        scope,
        re.IGNORECASE,
    )

    if not match:
        raise ValueError(
            f"Invalid group: {scope}"
        )

    group = match.group(1).upper()

    subgroup = (
        match.group(2).upper()
        if match.group(2)
        else None
    )

    return group, subgroup


def calendar_name(
    group,
    subgroup,
):

    if subgroup:
        return (
            f"ITECH 1 - {group} - {subgroup}"
        )

    return (
        f"ITECH 1 - {group}"
    )


# -----------------------------------------------------------------
# ICS GENERATION
# -----------------------------------------------------------------

def build_ics(
    events,
    group,
    subgroup,
    rooms_by_date,
):

    stamp = (
        "20260901T000000Z"
    )

    name = calendar_name(
        group,
        subgroup,
    )

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//planning-itech//FR",
        "CALSCALE:GREGORIAN",
        f"X-WR-CALNAME:{name}",
        f"X-WR-TIMEZONE:{TZ}",
        "REFRESH-INTERVAL;VALUE=DURATION:PT1H",
        "X-PUBLISHED-TTL:PT1H",
        VTZ,
    ]

    count_events = 0
    matched_rooms = 0
    courses_needing_rooms = 0

    for event in events:

        # ---------------------------------------------------------
        # GROUP FILTER
        # ---------------------------------------------------------

        if group not in event["groups"]:
            continue

        # ---------------------------------------------------------
        # A / B FILTER
        # ---------------------------------------------------------

        event_subgroups = (
            event.get(
                "subgroups",
                {},
            ).get(
                group,
                set(),
            )
        )

        if subgroup is not None:

            if subgroup not in event_subgroups:
                continue

        # ---------------------------------------------------------
        # TEXT
        # ---------------------------------------------------------

        event_lines = [
            line.strip()
            for line in event["text"].split("\n")
            if line.strip()
        ]

        title = (
            re.sub(
                r"\s+",
                " ",
                event_lines[0],
            )
            if event_lines
            else "Cours"
        )

        description = [
            re.sub(
                r"\s+",
                " ",
                line,
            )
            for line in event_lines[1:]
        ]

        # ---------------------------------------------------------
        # A / B INFORMATION
        #
        # ONLY the global calendar receives the A/B information.
        #
        # G2:
        #     Cours A -> "Groupe A"
        #     Cours B -> "Groupe B"
        #     Cours A+B -> "Groupes A et B"
        #
        # The A/B calendars themselves don't need this line.
        # ---------------------------------------------------------

        if subgroup is None:

            if event_subgroups == {"A"}:

                description.append(
                    "Groupe A"
                )

            elif event_subgroups == {"B"}:

                description.append(
                    "Groupe B"
                )

            elif event_subgroups == {
                "A",
                "B",
            }:

                description.append(
                    "Groupes A et B"
                )

        # ---------------------------------------------------------
        # ROOMS
        # ---------------------------------------------------------

        rooms = find_rooms(
            event,
            group,
            rooms_by_date,
        )

        location = ""

        if not event["allday"]:

            if rooms:

                location = " / ".join(
                    room["room"]
                    for room in rooms
                )

                if len(rooms) > 1:

                    description.append(
                        "Salles : "
                        + " ; ".join(
                            f"{room['room']}"
                            + (
                                f" ({room['teacher']})"
                                if room["teacher"]
                                else ""
                            )
                            for room in rooms
                        )
                    )

            if (
                not norm(
                    title
                ).startswith(
                    "e learning"
                )
                and not re.match(
                    r"\d{1,2}h",
                    title,
                )
            ):

                courses_needing_rooms += 1

                if rooms:
                    matched_rooms += 1

        # ---------------------------------------------------------
        # UID
        # ---------------------------------------------------------

        scope = (
            f"{group}-{subgroup}"
            if subgroup
            else group
        )

        key = (
            f"{scope}|"
            f"{event['date']}|"
            f"{event.get('start', 'allday')}|"
            f"{event.get('end', '')}|"
            f"{norm(title)}"
        )

        uid = (
            hashlib.sha1(
                key.encode(
                    "utf-8"
                )
            ).hexdigest()[:20]
            + "@planning-itech"
        )

        # ---------------------------------------------------------
        # VEVENT
        # ---------------------------------------------------------

        lines.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{stamp}",
                f"SUMMARY:{esc(title)}",
            ]
        )

        # ---------------------------------------------------------
        # ALL DAY
        # ---------------------------------------------------------

        if event["allday"]:

            lines.extend(
                [
                    (
                        "DTSTART;VALUE=DATE:"
                        f"{event['date']:%Y%m%d}"
                    ),
                    (
                        "DTEND;VALUE=DATE:"
                        f"{event['end_date']:%Y%m%d}"
                    ),
                    "TRANSP:TRANSPARENT",
                ]
            )

        # ---------------------------------------------------------
        # TIMED EVENT
        # ---------------------------------------------------------

        else:

            start = event["start"]
            end = event["end"]

            lines.extend(
                [
                    (
                        f"DTSTART;TZID={TZ}:"
                        f"{event['date']:%Y%m%d}"
                        f"T{start // 60:02d}"
                        f"{start % 60:02d}00"
                    ),
                    (
                        f"DTEND;TZID={TZ}:"
                        f"{event['date']:%Y%m%d}"
                        f"T{end // 60:02d}"
                        f"{end % 60:02d}00"
                    ),
                ]
            )

        # ---------------------------------------------------------
        # LOCATION
        # ---------------------------------------------------------

        if location:

            lines.append(
                f"LOCATION:{esc(location)}"
            )

        # ---------------------------------------------------------
        # DESCRIPTION
        # ---------------------------------------------------------

        if description:

            lines.append(
                "DESCRIPTION:"
                + esc(
                    chr(10).join(
                        description
                    )
                )
            )

        lines.append(
            "END:VEVENT"
        )

        count_events += 1

    lines.append(
        "END:VCALENDAR"
    )

    output = (
        "\r\n".join(
            fold(line)
            if not line.startswith(
                "BEGIN:VTIMEZONE"
            )
            else line
            for line in lines
        )
        + "\r\n"
    )

    return (
        output,
        count_events,
        matched_rooms,
        courses_needing_rooms,
    )


# -----------------------------------------------------------------
# GOOGLE SHEETS DOWNLOAD
# -----------------------------------------------------------------

def download(
    sheet_id,
    destination,
):

    url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{sheet_id}/export?format=xlsx"
    )

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0"
        },
    )

    try:

        data = urllib.request.urlopen(
            request,
            timeout=120,
        ).read()

    except Exception as error:

        raise SystemExit(
            "Download failed "
            f"({url}): {error}"
        )

    if not data.startswith(
        b"PK"
    ):

        raise SystemExit(
            "Download failed "
            f"({url}): the Google Sheet "
            "does not appear to be public."
        )

    with open(
        destination,
        "wb",
    ) as file:

        file.write(data)


# -----------------------------------------------------------------
# OUTPUT FILENAME
# -----------------------------------------------------------------

def output_filename(
    group,
    subgroup,
):

    if subgroup:

        return (
            f"itech1_{group}_{subgroup}.ics"
        )

    return (
        f"itech1_{group}.ics"
    )


# -----------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------

def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "planning",
        nargs="?",
    )

    parser.add_argument(
        "salles",
        nargs="?",
    )

    parser.add_argument(
        "--download",
        action="store_true",
        help=(
            "Download the public Google Sheets"
        ),
    )

    parser.add_argument(
        "--groupe",
        choices=[
            "G1",
            "G1-A",
            "G1-B",
            "G2",
            "G2-A",
            "G2-B",
            "G3",
            "G3-A",
            "G3-B",
        ],
    )

    parser.add_argument(
        "--tous",
        action="store_true",
        help=(
            "Generate all global and A/B calendars"
        ),
    )

    parser.add_argument(
        "-o",
        "--out",
    )

    parser.add_argument(
        "--outdir",
        default=".",
    )

    args = parser.parse_args()

    # -------------------------------------------------------------
    # DOWNLOAD
    # -------------------------------------------------------------

    if args.download:

        args.planning = "planning.xlsx"
        args.salles = "salles.xlsx"

        print(
            "Downloading planning..."
        )

        download(
            PLANNING_ID,
            args.planning,
        )

        print(
            "Downloading rooms..."
        )

        download(
            SALLES_ID,
            args.salles,
        )

    # -------------------------------------------------------------
    # VALIDATION
    # -------------------------------------------------------------

    if not (
        args.planning
        and args.salles
    ):

        parser.error(
            "Provide PLANNING.xlsx and SALLES.xlsx, "
            "or use --download."
        )

    if args.out and args.tous:

        parser.error(
            "--out cannot be used with --tous."
        )

    # -------------------------------------------------------------
    # READ PLANNING
    # -------------------------------------------------------------

    print(
        "Reading planning..."
    )

    events = read_planning(
        args.planning
    )

    if len(events) < MIN_EVENTS:

        raise SystemExit(
            f"Only {len(events)} events were read. "
            "The spreadsheet structure may have changed. "
            "Nothing was overwritten."
        )

    print(
        f"Planning: {len(events)} events detected."
    )

    # -------------------------------------------------------------
    # READ ROOMS
    # -------------------------------------------------------------

    print(
        "Reading rooms..."
    )

    rooms = read_rooms(
        args.salles
    )

    rooms_by_date = defaultdict(
        list
    )

    for room in rooms:

        rooms_by_date[
            room["date"]
        ].append(
            room
        )

    print(
        f"Rooms: {len(rooms)} entries detected."
    )

    # -------------------------------------------------------------
    # OUTPUT DIRECTORY
    # -------------------------------------------------------------

    os.makedirs(
        args.outdir,
        exist_ok=True,
    )

    # -------------------------------------------------------------
    # CALENDARS
    # -------------------------------------------------------------

    if args.tous:

        calendars = [
            ("G1", None),
            ("G1", "A"),
            ("G1", "B"),

            ("G2", None),
            ("G2", "A"),
            ("G2", "B"),

            ("G3", None),
            ("G3", "A"),
            ("G3", "B"),
        ]

    else:

        scope = (
            args.groupe
            or "G1"
        )

        group, subgroup = (
            parse_calendar_scope(
                scope
            )
        )

        calendars = [
            (
                group,
                subgroup,
            )
        ]

    # -------------------------------------------------------------
    # GENERATE
    # -------------------------------------------------------------

    generated = 0

    for group, subgroup in calendars:

        ics, count, matched, need = (
            build_ics(
                events,
                group,
                subgroup,
                rooms_by_date,
            )
        )

        if (
            args.out
            and not args.tous
        ):

            output_path = args.out

        else:

            output_path = os.path.join(
                args.outdir,
                output_filename(
                    group,
                    subgroup,
                ),
            )

        with open(
            output_path,
            "w",
            encoding="utf-8",
            newline="",
        ) as file:

            file.write(
                ics
            )

        print(
            f"{output_path} : "
            f"{count} events, "
            f"rooms found for "
            f"{matched}/{need} courses"
        )

        generated += 1

    # -------------------------------------------------------------
    # LAST UPDATE
    # -------------------------------------------------------------

    with open(
        os.path.join(
            args.outdir,
            "last_update.txt",
        ),
        "w",
        encoding="utf-8",
    ) as file:

        file.write(
            dt.date.today().isoformat()
            + "\n"
        )

    print(
        f"Generated {generated} calendar(s)."
    )


# -----------------------------------------------------------------
# ENTRY POINT
# -----------------------------------------------------------------

if __name__ == "__main__":
    main()
