# STATS EXTENSIONS REPORT command

__author__ = "JKP"
__version__ = "1.0.0"

# history
# 07-apr-2023  Original version


import spss, spssdata
from extension import Template, Syntax, processcmd

import glob, random, zipfile, os.path, re
from collections import namedtuple

# debugging
#try:
    #import wingdbstub
    #import threading
    #wingdbstub.Ensure()
    #wingdbstub.debugger.SetDebugThreads({threading.get_ident(): 1})
#except:
    #pass

# TODO: two types
extinfo = namedtuple("ext", ["file", "display_Name", "order", "version", "type", "Plugins",
    "Python_Modules", "R_Packages", "loc", "summary"])
manifestnames = {"Summary: ":"summary", "Version: ":"version", "Plugins: ":"Plugins",
    "R-Packages: ":"R_Packages", "Python-Modules: ":"Python_Modules", "Display-Name: ": "display_Name"}

# do STATS EXTENSION REPORT command
def dorpt(nfilter=""):
    
    # The CDB generates code with unexpected trailing blank on the pattern
    nfilter = nfilter.rstrip()
    nfilterc = re.compile(nfilter, re.I)
    # get extension info
    dsname = spss.ActiveDataset()
    if  dsname == "*":
        tempname = "D" + str(random.uniform(.05, 1))
        spss.Submit(f"DATASET NAME {tempname}.")
    else:
        tempname = None
    extds = "D" + str(random.uniform(.05, 1))
    tag = "T" + str(random.uniform(.05, 1))
    spss.Submit(f"""DATASET DECLARE {extds}.
OMS SELECT TABLES /IF SUBTYPES=['System Settings']
/DESTINATION OUTFILE="{extds}" FORMAT=SAV VIEWER=NO
/TAG = {tag}.""")
    # The row labels in this table are not currently translated,
    # but just in case that changes, forcing English output.
    spss.Submit(f"""PRESERVE.
SET OLANG=ENGLISH.
SHOW EXT.
RESTORE.
OMSEND TAG="{tag}".""") 
    spss.Submit(f"""DATASET ACTIVATE {extds}.""")

    # label (3) EXTPATHS CDIALOGS OR EXTGPATHS EXTENSIONS
    # number (4) search order 
    # setting (6) location
    curs = spssdata.Spssdata(indexes=[3, 4, 6])
    tbl = curs.fetchall()
    curs.CClose()
    spss.Submit(f"DATASET CLOSE {extds}")
    tlen = len(tbl)
    
    # location for new-style extensions does not appear in SHOW EXT output, so add a guess
    for i in range(tlen):
        if tbl[i][0].rstrip() == "EXTPATHS EXTENSIONS":
            parts = os.path.split(tbl[i][2])
            if parts[-1].rstrip().lower() == "extensions":
                newextloc = parts[0] + os.path.sep + "xtensions"
                tbl.append(["EXTPATHS EXTENSIONS", "1", newextloc])
    # if original active dataset was empty, it might be gone
    if spss.GetDatasets():        
        if tempname is None:
            spss.Submit(f"DATASET ACTIVATE {dsname}.")
        else:
            spss.Submit(f"""DATASET ACTIVATE {tempname}.
            DATASET CLOSE {tempname}.""")
    
    ptdata = []
    for i in range(len(tbl)):
        loc = tbl[i][-1].rstrip()
        # glob is not case sensitive, at least on Windows
        files = glob.iglob("./*/*.sp*", root_dir=loc)  

        for f in files:
            if not os.path.splitext(f)[-1].lower() in [".spe", ".spxt"]:
                continue
            if nfilter and re.search(nfilterc, os.path.split(f)[-1]) is None:
                continue
            d = {}
            d['file'] = os.path.basename(f)
            d['loc'] = loc + os.path.dirname(f)[1:]
            d['order'] = int(tbl[i][1])
            d['type'] = "spe"

            with zipfile.ZipFile(loc + f[1:]) as zf:
                # case correct the name
                contents = zf.namelist()
                # zip extractor is case sensitive :-(
                for c in contents:
                    if c.startswith("META-INF"):
                        metaname = "META-INF" + "/" + os.path.split(c)[-1]
                        with zf.open(metaname) as mm:
                            for line in mm:
                                sline = line.decode()    # zip read gives byte objects.  utf-8 asssumed
                                for tag in manifestnames:
                                    if sline.startswith(tag):
                                        tag.replace("-", "_")
                                        d[manifestnames[tag]] = sline[len(tag):]
            if not "Python_Modules" in d:
                d["Python_Modules"] = "None"
            if not "R_Packages" in d:
                d["R_Packages"] = "None"
            if not "Plugins" in d:
                d["Plugins"] = "None"
            if not "display_Name" in d:
                d["display_Name"] = d["file"].replace("_", " ")
            ptdata.append((extinfo(**d)))       
        
    # make pivot table
    if not ptdata:
        print(_("""No extensions were found."""))
        return
    ptdata.sort(key=lambda s: [s[0].lower(), s[1].lower(), s[2], s[4].lower(), s[8].lower()])
    spss.StartProcedure("Extension Report")
    pt = spss.BasePivotTable(_("Installed Extension Commands"), "Extensioncmds")
    if nfilter:
        pt.TitleFootnotes(f"""Filtered by {nfilter}""")

    pt.SimplePivotTable(rowlabels = [str(i+1) for i in range(len(ptdata))],
        collabels=extinfo._fields,
        cells=ptdata)
    
    spss.EndProcedure()

def  Run(args):
    """Execute the STATS PREPROCESS command"""

    args = args[list(args.keys())[0]]

    oobj = Syntax([
        Template("FILTER", subc="",  ktype="str", var="nfilter", islist=False)])
        
        
    #enable localization
    global _
    try:
        _("---")
    except:
        def _(msg):
            return msg

    # A HELP subcommand overrides all else
    if "HELP" in args:
        #print helptext
        helper()
    else:
        processcmd(oobj, args, dorpt)

def helper():
    """open html help in default browser window
    
    The location is computed from the current module name"""
    
    import webbrowser, os.path
    
    path = os.path.splitext(__file__)[0]
    helpspec = "file://" + path + os.path.sep + \
         "markdown.html"
    
    # webbrowser.open seems not to work well
    browser = webbrowser.get()
    if not browser.open_new(helpspec):
        print(("Help file not found:" + helpspec))
try:    #override
    from extension import helper
except:
    pass        
