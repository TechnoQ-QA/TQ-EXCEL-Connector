"""Techno Q only: prepares the NetSuite package (netsuite/sdf), the SuiteCloud project that installs everything
the connector needs in a customer's NetSuite account in one go.

    python scripts/package_netsuite.py

copies the three scripts from netsuite/ into the package's File Cabinet folder and checks the package:
- every object file is well-formed XML and points at a script that is in the package;
- the Suitelet's script and deployment IDs are the ones the RESTlet calls;
- the Suitelet runs as Administrator and is not available without login.
Then deploy it with the SuiteCloud CLI (see netsuite/sdf/README.md).
"""

import os
import re
import shutil
import sys
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCE = os.path.join(ROOT, "netsuite")
PROJECT = os.path.join(SOURCE, "sdf", "src")
FOLDER = "/SuiteScripts/TQ Excel Connector"
SCRIPTS = ("tq_excel_core.js", "tq_excel_restlet.js", "tq_excel_elevated.js")


def copy():
    target = os.path.join(PROJECT, "FileCabinet", *FOLDER.strip("/").split("/"))
    os.makedirs(target, exist_ok=True)
    for name in SCRIPTS:
        shutil.copyfile(os.path.join(SOURCE, name), os.path.join(target, name))
    return target


def check():
    """A list of problems; empty when the package is complete."""
    problems = []
    for name in ("manifest.xml", "deploy.xml"):
        try:
            ET.parse(os.path.join(PROJECT, name))
        except (ET.ParseError, OSError) as err:
            problems.append("%s: %s" % (name, err))
    objects = os.path.join(PROJECT, "Objects")
    found = {}
    for name in sorted(os.listdir(objects)):
        try:
            root = ET.parse(os.path.join(objects, name)).getroot()
        except ET.ParseError as err:
            problems.append("%s: %s" % (name, err))
            continue
        if root.get("scriptid") + ".xml" != name:
            problems.append("%s: its scriptid is %s" % (name, root.get("scriptid")))
        found[root.get("scriptid")] = root
        script = root.findtext("scriptfile")
        if script:
            path = script.strip("[]")
            if not path.startswith(FOLDER + "/") or os.path.basename(path) not in SCRIPTS:
                problems.append("%s: scriptfile %s is not one of the package's scripts" % (name, path))
            elif not os.path.exists(os.path.join(PROJECT, "FileCabinet", *path.strip("/").split("/"))):
                problems.append("%s: %s is missing; run this script to copy it" % (name, path))
    suitelet = found.get("customscript_tq_excel_elevated")
    restlet_code = open(os.path.join(SOURCE, "tq_excel_restlet.js"), encoding="utf-8").read()
    if suitelet is None:
        problems.append("The elevated Suitelet object is missing")
    else:
        deployment = suitelet.find("scriptdeployments/scriptdeployment")
        if "scriptId: '%s'" % suitelet.get("scriptid") not in restlet_code or \
                "deploymentId: '%s'" % deployment.get("scriptid") not in restlet_code:
            problems.append("The RESTlet calls a Suitelet with other IDs than the package installs")
        if deployment.findtext("runasrole") != "ADMINISTRATOR":
            problems.append("The elevated Suitelet must run as Administrator")
        if deployment.findtext("isonline") != "F":
            problems.append("The elevated Suitelet must not be available without login")
    for name in SCRIPTS:
        packed = os.path.join(PROJECT, "FileCabinet", *FOLDER.strip("/").split("/"), name)
        if os.path.exists(packed) and open(packed, "rb").read() != open(os.path.join(SOURCE, name), "rb").read():
            problems.append("%s in the package differs from netsuite/%s; run this script" % (name, name))
    version = re.search(r"const VERSION = (\d+)", open(os.path.join(SOURCE, "tq_excel_core.js"), encoding="utf-8").read())
    if not version:
        problems.append("tq_excel_core.js has no VERSION")
    return problems


if __name__ == "__main__":
    where = copy()
    issues = check()
    for issue in issues:
        print("PROBLEM:", issue)
    if issues:
        sys.exit(1)
    print("Package ready in %s (scripts copied to %s)." % (os.path.dirname(PROJECT), where))
