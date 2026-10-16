"""
Canonical 32-council list and name resolver.

Kept intentionally identical to the COUNCILS / ALIASES tables in
krisis-verification/scripts/normalise_workbook.py. If that list ever changes,
copy the change here too -- there is no runtime dependency between the two
skills, on purpose, so this one still works if it is installed on its own.
"""
import re
from collections import OrderedDict

COUNCILS = [
    ("S12000033", "Aberdeen City"),
    ("S12000034", "Aberdeenshire"),
    ("S12000041", "Angus"),
    ("S12000035", "Argyll and Bute"),
    ("S12000036", "City of Edinburgh"),
    ("S12000005", "Clackmannanshire"),
    ("S12000013", "Na h-Eileanan Siar"),
    ("S12000006", "Dumfries and Galloway"),
    ("S12000042", "Dundee City"),
    ("S12000008", "East Ayrshire"),
    ("S12000045", "East Dunbartonshire"),
    ("S12000010", "East Lothian"),
    ("S12000011", "East Renfrewshire"),
    ("S12000014", "Falkirk"),
    ("S12000047", "Fife"),
    ("S12000049", "Glasgow City"),
    ("S12000017", "Highland"),
    ("S12000018", "Inverclyde"),
    ("S12000019", "Midlothian"),
    ("S12000020", "Moray"),
    ("S12000021", "North Ayrshire"),
    ("S12000050", "North Lanarkshire"),
    ("S12000023", "Orkney Islands"),
    ("S12000048", "Perth and Kinross"),
    ("S12000038", "Renfrewshire"),
    ("S12000026", "Scottish Borders"),
    ("S12000027", "Shetland Islands"),
    ("S12000028", "South Ayrshire"),
    ("S12000029", "South Lanarkshire"),
    ("S12000030", "Stirling"),
    ("S12000039", "West Dunbartonshire"),
    ("S12000040", "West Lothian"),
]
NAME_BY_CODE = OrderedDict(COUNCILS)
CODE_BY_NAME = {v: k for k, v in COUNCILS}

ALIASES = {
    "aberdeen": "Aberdeen City",
    "aberdeen city council": "Aberdeen City",
    "aberdeenshire council": "Aberdeenshire",
    "angus": "Angus", "angus council": "Angus",
    "argyle & bute": "Argyll and Bute", "argyll and bute council": "Argyll and Bute",
    "argyll & bute": "Argyll and Bute", "argyll & bute council": "Argyll and Bute",
    "city of edinburgh": "City of Edinburgh", "city of edinburgh council": "City of Edinburgh",
    "edinburgh": "City of Edinburgh",
    "clackmannanshire": "Clackmannanshire", "clackmannanshire council": "Clackmannanshire",
    "comhairle nan eilean siar": "Na h-Eileanan Siar",
    "western isles": "Na h-Eileanan Siar", "outer hebrides": "Na h-Eileanan Siar",
    "dumfries & galloway": "Dumfries and Galloway",
    "dumfries and galloway council": "Dumfries and Galloway",
    "dundee": "Dundee City", "dundee city council": "Dundee City",
    "east ayrshire": "East Ayrshire", "east ayrshire council": "East Ayrshire",
    "east dumbartonshire": "East Dunbartonshire",
    "east dunbartonshire council": "East Dunbartonshire",
    "east lothian": "East Lothian", "east lothian council": "East Lothian",
    "east renfrewshire": "East Renfrewshire", "east renfrewshire council": "East Renfrewshire",
    "falkirk": "Falkirk", "falkirk council": "Falkirk",
    "fife": "Fife", "fife council": "Fife",
    "glasgow city": "Glasgow City", "glasgow city council": "Glasgow City",
    "glasgow": "Glasgow City",
    "highland": "Highland", "highland council": "Highland",
    "highlands": "Highland",
    "inverclyde": "Inverclyde", "inverclyde council": "Inverclyde",
    "midlothian": "Midlothian", "midlothian council": "Midlothian",
    "moray": "Moray", "moray council": "Moray",
    "north ayrshire": "North Ayrshire", "north ayrshire council": "North Ayrshire",
    "north lanarkshire": "North Lanarkshire", "north lanarkshire council": "North Lanarkshire",
    "orkney": "Orkney Islands", "orkney islands council": "Orkney Islands",
    "perth & kinross": "Perth and Kinross", "perth and kinross council": "Perth and Kinross",
    "renfrewshire": "Renfrewshire", "renfrewshire council": "Renfrewshire",
    "scottish borders": "Scottish Borders", "scottish borders council": "Scottish Borders",
    "borders": "Scottish Borders",
    "shetland": "Shetland Islands", "shetland islands council": "Shetland Islands",
    "south ayrshire": "South Ayrshire", "south ayrshire council": "South Ayrshire",
    "south lanarkshire": "South Lanarkshire", "south lanarkshire council": "South Lanarkshire",
    "stirling": "Stirling", "stirling council": "Stirling",
    "west dunbarton": "West Dunbartonshire",
    "west dunbartonshire council": "West Dunbartonshire",
    "west lothian": "West Lothian", "west lothian council": "West Lothian",
}


def canon_la(raw):
    """Return one of the 32 canonical council names, or None if unrecognised.
    Never guesses a nearest match -- an unrecognised header is left alone so
    the caller can flag it rather than silently mis-file records under it."""
    if raw is None:
        return None
    key = re.sub(r"\s+", " ", str(raw)).strip().lower()
    # exact canonical name (case-insensitive)
    for _, name in COUNCILS:
        if name.lower() == key:
            return name
    return ALIASES.get(key)
