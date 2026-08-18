"""
Mapping layer: form answer -> reporting category -> Excel column / value.

This is the single source of truth for how demographic answers are stored and
how they are written into the existing Excel reporting workbook.

Coding rules were determined by inspecting the workbook ("Feb. Numbers
Final(1).xlsx"). Findings:

  * Every demographic / category column is coded ``1`` when it applies and is
    left BLANK when it does not. There is NO ``Yes=1 / No=2`` scheme.
  * ``House Size`` holds an actual integer count.
  * The large numbers seen in the sheets (39, 63, 92 ...) are ``=SUM()`` totals
    rows, not per-client data.
  * Race, Household-type and Population-category behave as single-select groups
    (exactly one ``1`` per client).
  * Header spellings in the workbook are idiosyncratic and are preserved
    verbatim below ("Veteren", "Two Parent " with a trailing space, etc.).

The canonical column order follows the newest / cleanest sheet in the
workbook, "June Complete", which is also the only sheet carrying a UID column.
"""

from __future__ import annotations

from datetime import date
from typing import Optional


# ---------------------------------------------------------------------------
# Single-select answer domains (value stored in the DB -> reporting category)
# ---------------------------------------------------------------------------

RACE_CHOICES = [
    ("white", "White"),
    ("black", "Black"),
    ("asian", "Asian"),
    ("american_indian", "American Indian"),
    ("other_multi", "Other / Multi-racial"),
]

# race code -> exact workbook header that should receive a 1
RACE_TO_HEADER = {
    "white": "White",
    "black": "Black",
    "asian": "Asian",
    "american_indian": "American India",
    "other_multi": "Other/Multi",
}

HOUSEHOLD_TYPE_CHOICES = [
    ("single_non_elderly", "Single, Non-Elderly"),
    ("elderly", "Elderly (62+)"),
    ("single_parent", "Single Parent"),
    ("two_parent", "Two Parent"),
    ("other", "Other"),
]

HOUSEHOLD_TYPE_TO_HEADER = {
    "single_non_elderly": "Single Non Elderly",
    "elderly": "Elderly (62+)",
    "single_parent": "Single Parent",
    "two_parent": "Two Parent ",  # trailing space is intentional (matches workbook)
    "other": "Other",
}

POPULATION_CHOICES = [
    ("adult", "Adult"),
    ("tay", "TAY (Transition-Age Youth, 18-24)"),
    ("family_with_minor", "Family With Minor"),
    ("senior", "Senior"),
]

POPULATION_TO_HEADER = {
    "adult": "Adult",
    "tay": "TAY",
    "family_with_minor": "Fam. With Minor",
    "senior": "Senior",
}

# Simple yes/no flags: DB stores 1/0, workbook gets 1 or blank.
YES_NO_FLAGS = {
    "veteran": "Veteren",  # workbook misspelling preserved
    "hispanic": "Hispanic",
    "female_head": "Female Head of",
    "disabled": "Disabled",
}


# ---------------------------------------------------------------------------
# Canonical intake sheet layout (order matters -> column A, B, C ...)
# ---------------------------------------------------------------------------
# Each entry: (exact header text, semantic "kind")
#   kind "text"  -> written as-is (string)
#   kind "date"  -> written as a date
#   kind "int"   -> written as an integer
#   kind "flag:<field>"   -> 1 if that yes/no flag is set else blank
#   kind "race:<code>"    -> 1 if race == code else blank
#   kind "hh:<code>"      -> 1 if household_type == code else blank
#   kind "pop:<code>"     -> 1 if population_category == code else blank

INTAKE_COLUMNS = [
    ("Last Name", "text:last_name"),
    ("First Name", "text:first_name"),
    ("Birthdate", "date:date_of_birth"),
    ("UID (If applicable)", "text:uid"),
    ("Veteren", "flag:veteran"),
    ("Hispanic", "flag:hispanic"),
    ("White", "race:white"),
    ("Black", "race:black"),
    ("Asian", "race:asian"),
    ("American India", "race:american_indian"),
    ("Other/Multi", "race:other_multi"),
    ("House Size", "int:house_size"),
    ("Single Non Elderly", "hh:single_non_elderly"),
    ("Elderly (62+)", "hh:elderly"),
    ("Single Parent", "hh:single_parent"),
    ("Two Parent ", "hh:two_parent"),
    ("Other", "hh:other"),
    ("Adult", "pop:adult"),
    ("TAY", "pop:tay"),
    ("Fam. With Minor", "pop:family_with_minor"),
    ("Senior", "pop:senior"),
    ("Female Head of", "flag:female_head"),
    ("Disabled", "flag:disabled"),
]

# Headers that are demographic/numeric and should be summed in the totals row.
TOTALS_START_HEADER = "Veteren"  # first summed column
TOTALS_ROW_LABEL = "TOTALS"

# ---------------------------------------------------------------------------
# Canonical Sign-In sheet layout
# ---------------------------------------------------------------------------
SIGNIN_COLUMNS = [
    ("Date", "date:visit_date"),
    ("Time", "text:visit_time"),
    ("Last Name", "text:last_name"),
    ("First Name", "text:first_name"),
    ("DOB", "date:date_of_birth"),
    ("Visitor Type", "visitor:visitor_type"),
    ("Services", "text:services"),
]

VISITOR_TYPE_CHOICES = [
    ("resident", "Resident (Lot O / Program Participant)"),
    ("walk_in", "Walk-In Visitor / Community Guest"),
]

VISITOR_TYPE_LABELS = dict(VISITOR_TYPE_CHOICES)

DEFAULT_SERVICES = "Employment; Work Force"


# ---------------------------------------------------------------------------
# Helpers to resolve a column "kind" to an actual Excel cell value.
# ---------------------------------------------------------------------------

def _flag_cell(value) -> Optional[int]:
    """Yes/no flag -> 1 when applicable, blank (None) otherwise."""
    return 1 if value else None


def intake_cell_value(kind: str, intake: dict, client: dict):
    """Resolve one intake column ``kind`` to the value written into Excel.

    ``intake`` and ``client`` are plain dicts of stored values.
    """
    op, _, arg = kind.partition(":")

    if op == "text":
        source = client if arg in client else intake
        return source.get(arg) or None
    if op == "date":
        source = client if arg in client else intake
        return source.get(arg) or None
    if op == "int":
        val = intake.get(arg)
        try:
            return int(val) if val not in (None, "") else None
        except (TypeError, ValueError):
            return None
    if op == "flag":
        return _flag_cell(intake.get(arg))
    if op == "race":
        return 1 if intake.get("race") == arg else None
    if op == "hh":
        return 1 if intake.get("household_type") == arg else None
    if op == "pop":
        return 1 if intake.get("population_category") == arg else None
    return None


def signin_cell_value(kind: str, visit: dict, client: dict):
    """Resolve one sign-in column ``kind`` to the value written into Excel."""
    op, _, arg = kind.partition(":")
    if op in ("text", "date"):
        source = client if arg in client else visit
        return source.get(arg) or None
    if op == "visitor":
        return VISITOR_TYPE_LABELS.get(visit.get(arg), visit.get(arg)) or None
    return None


def suggest_population_category(dob: Optional[date]) -> str:
    """Default the population-category radio from age. Staff/client may change it."""
    if not dob:
        return "adult"
    today = date.today()
    age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    if age < 25:
        return "tay"
    if age >= 62:
        return "senior"
    return "adult"


def suggest_household_type(dob: Optional[date]) -> str:
    """Default household type from age (elderly 62+ else single non-elderly)."""
    if not dob:
        return "single_non_elderly"
    today = date.today()
    age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    return "elderly" if age >= 62 else "single_non_elderly"
