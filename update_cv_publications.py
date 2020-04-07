import re

from pybtex.database.input import bibtex
from pybtex.style.formatting import BaseStyle
import json

parser = bibtex.Parser()
bib_data = parser.parse_file("publications.bib")

pubs = []

from pybtex.style.names.plain import NameStyle
from pybtex.style.formatting import BaseStyle, toplevel
from pybtex.style.template import field, join, optional, sentence, words
from pybtex.style.sorting import BaseSortingStyle


for key in bib_data.entries:
    entry = bib_data.entries[key]
    fields = entry.fields
    desc = {}
    if "wwwhidden" in fields and fields["wwwhidden"]:
        continue
    desc["type"] = fields["wwwtype"]
    if desc["type"] == "working":
        continue
    desc["name"] = fields["title"]
    desc["authors"] = ""
    for person in entry.persons["author"]:
        nbsp = " "
        name = NameStyle().format(person, abbr=True).format().render_as("text")
        name = name.replace(". ", "."+nbsp)
        desc["authors"] += name + ", "
    desc["authors"] = desc["authors"][:-2]
    if "location" in fields:
        desc["location"] = fields["location"]
    if desc["type"] == "journal":
        desc["publisher"] = fields["journal"]
    elif desc["type"] == "working":
        # Preprint
        pass
    else:
        desc["publisher"] = fields["booktitle"]
    desc["publisher"] = desc["publisher"].replace("The Journal", "Journal")
    desc["publisher"] = desc["publisher"].replace("Proceedings", "Proc.")
    desc["publisher"] = desc["publisher"].replace("International", "Int.")
    desc["publisher"] = desc["publisher"].replace("Conference", "Conf.")
    desc["publisher"] = desc["publisher"].replace("Symposium", "Symp.")
    desc["publisher"] = desc["publisher"].replace(" and ", " ")
    desc["publisher"] = desc["publisher"].replace(" on ", " ")
    desc["publisher"] = desc["publisher"].replace(" of the ", " ")
    desc["publisher"] = re.sub(r' \([^)]*\)', '', desc["publisher"])
    desc["releaseDate"] = fields["month"] + " " + fields["year"]
    desc["link"] = "https://nickwalker.us/publications/" + key
    pubs.append(desc)

with open("cv.json", 'r') as f:
    all_data = json.load(f)

all_data["publications"] = pubs

with open("cv.json", 'w') as f:
    json.dump(all_data, f, indent=True)
