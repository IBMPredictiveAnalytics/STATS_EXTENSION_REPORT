# STATS EXTENSIONS REPORT command

__author__ = "JKP"
__version__ = "1.1.0"

# history
# 07-apr-2023  Original version
# 12-mar-2024  New options for updates and uninstalled list


import spss, spssdata
from extension import Template, Syntax, processcmd

import glob, random, zipfile, os.path, re, copy, json
from collections import namedtuple

# debugging
try:
    import wingdbstub
    import threading
    wingdbstub.Ensure()
    wingdbstub.debugger.SetDebugThreads({threading.get_ident(): 1})
except:
    pass

# TODO: two types
extinfo = namedtuple("ext", ["file", "display_Name", "order", "version", "type", "Plugins",
    "Python_Modules", "R_Packages", "loc", "summary"])
manifestnames = {"Summary: ":"summary", "Version: ":"version", "Plugins: ":"Plugins",
    "R-Packages: ":"R_Packages", "Python-Modules: ":"Python_Modules", "Display-Name: ": "display_Name"}
nt = namedtuple("extdata", ["filename", "version", "summary", "productversion"])

# do STATS EXTENSION REPORT command
def dorpt(nfilter="", updaterpt=True, uninstalledrpt=False):
    
    # The CDB generates code with unexpected trailing blank on the pattern
    nfilter = nfilter.rstrip()
    nfilterc = re.compile(nfilter, re.I)
    ptdata, ptdatafiltered= getinstalledpackages(nfilter, nfilterc)
    if nfilter == "":
        using = ptdata
    else:
        using = ptdatafiltered
    if updaterpt or uninstalledrpt:
        ghdata = Extinfo()
        
    if not ptdata or (nfilter != "" and not ptdatafiltered):
        print(_("""No extensions were found."""))
        if not uninstalledrpt:
            return
    else:
        if updaterpt:
            for i, f in enumerate(using):
                f = os.path.splitext(f.file)[0].lower()
                if not f in ghdata.exts:
                    f = f.replace(" ", "_").lower()
                if f in ghdata.exts and versionLT(using[i].version, ghdata.exts[f].version):
                    using[i] = using[i]._replace(file=using[i].file + "*")  # sorry
                
        # order, s[2] has type CellText.Number, which has no comparison method, so convert (painfully) to an int for sorting
        ptdata.sort(key=lambda s: [s[0].lower(), s[1].lower(), s[2].data['value'], s[4].lower(), s[8].lower()])
        if nfilter != "":
            ptdatafiltered.sort(key=lambda s: [s[0].lower(), s[1].lower(), s[2].data['value'], s[4].lower(), s[8].lower()])    
    
        spss.StartProcedure("Extension Report")
        if updaterpt:
            caption = _("""* Update available from the Extension Hub""")
        else:
            caption = _("""Extensions not checked for updates""")
        pt = spss.BasePivotTable(_("Installed Extension Commands") , "Extensioncmds", caption = caption)
        if nfilter:
            pt.TitleFootnotes(f"""Filtered by {nfilter}""")
        if nfilter == "":
            pt.SimplePivotTable(rowlabels = [str(i+1) for i in range(len(ptdata))],
            collabels=extinfo._fields,
            cells=ptdata)
        else:
            pt.SimplePivotTable(rowlabels = [str(i+1) for i in range(len(ptdatafiltered))],
            collabels=extinfo._fields,
            cells=ptdatafiltered)
        pt.SetDefaultFormatSpec(6)
        
    if uninstalledrpt:
        douninstalledreport(ptdata, ghdata)
    
    spss.EndProcedure()

def douninstalledreport(ptdata, ghdata):
    """display table of extensions not install
    
    ptdata is the entire collections of installed exstensions
    ghdata is the collection of extensions online"""
    
    for f in ptdata:
        f = os.path.splitext(f.file)[0].lower()
        f = re.sub(r"\*$", "", f)  # remove any update available flag
        if f in ghdata.exts:
            del(ghdata.exts[f])
        else:
            f = f.replace(" ", "_")
            if f in ghdata.exts:
                del(ghdata.exts[f])
    tabledata = [(item.filename, item.summary, item.productversion) for item in ghdata.exts.values()]
    tabledata.sort()
    ptu = spss.BasePivotTable(_("""Uninstalled Extension Commands"""), "uninstalledexts")
    ptu.SimplePivotTable(rowlabels = [str(i+1) for i in range(len(ghdata.exts))], collabels = [_("File"), _("Summary"),
        _("Minimum SPSS Version")], cells=tabledata)

def getinstalledpackages(nfilter, nfilterc):
    # get extension info

    tag = "T" + str(random.uniform(.05, 1))
    extpathsxpaths = """//pivotTable//group[@text="EXTPATHS EXTENSIONS"]//category[@text="Setting"]/cell/@*"""
    extpathsorder = """//pivotTable//group[@text="EXTPATHS EXTENSIONS"]/category//@number"""
    spss.Submit(f"""OMS SELECT TABLES /IF SUBTYPES=['System Settings']
/DESTINATION XMLWORKSPACE="{tag}" FORMAT=OXML VIEWER=NO
/TAG = {tag}.
PRESERVE.
SET OLANG=ENGLISH.
SHOW EXT.
RESTORE.
OMSEND TAG="{tag}".""")
    orders = spss.EvaluateXPath(tag, "/", extpathsorder)
    locs = spss.EvaluateXPath(tag, "/",  extpathsxpaths)
    spss.DeleteXPathHandle(tag)
    tbl = list(zip(orders, locs))
    tlen = len(tbl)
    for i in range(tlen):
        parts = os.path.split(tbl[i][-1])
        if parts[-1].rstrip().lower() == "extensions":
            newextloc = parts[0] + os.path.sep + "xtensions"
            tbl.append((1, newextloc))
            break
    
    ptdata = []
    ptdatafiltered = []
    for i in range(len(tbl)):
        loc = tbl[i][-1].rstrip()
        # glob is not case sensitive, at least on Windows
        files = glob.glob("./*/*.sp*", root_dir=loc)  

        for f in files:
            if not os.path.splitext(f)[-1].lower() in [".spe", ".spxt"]:
                continue
            d = {}
            d['file'] = os.path.basename(f)
            d['loc'] = loc + os.path.dirname(f)[1:]
            d['order'] = spss.CellText.Number(int(tbl[i][0]), spss.FormatSpec.Count)
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
                                sline = sline.replace("\n", "")  # zip line may include \n at the end
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
            if nfilter and re.search(nfilterc, os.path.split(f)[-1]) is not None:
                ptdatafiltered.append((extinfo(**d)))           
    return ptdata, ptdatafiltered
    

def versionLT(v1, v2):
    """return True if v1 < v2 elswe False
    v1 and v2 are string version numbers typically including dotted parts"""
    
    try:        
        v1 = [int(part) for part in v1.split(".")]
        v2 = [int(part) for part in v2.split(".")]
        return v1 < v2
    except:
        return False
    
import requests, io

class Extinfo():
    "Get Extension Hub information about extensions using the index file"
    def __init__(self):
        self.exts = {}
        index = requests.get("https://raw.githubusercontent.com/IBMPredictiveAnalytics/IBMPredictiveAnalytics.github.io/master/resbundles/statisitcs/extension_index_resbundles.zip")
        if index.status_code > 200:
            raise SystemError(f"SPSS extension repository access failed: {index.reason}")
        zfile = zipfile.ZipFile(io.BytesIO(index.content))
        ###tempdir = tempfile.mkdtemp()
        doc = json.load(zfile.open("extension_info_index.json"))
        doc =  doc['extension_index']
        for d in doc:
            exten = d['extension_detail_info']
            name = exten["Name"].replace(" ", "_").lower()
            # keys:Name (may have blanks), Display-Name, Command-Specs, Summary, Version, Product-Version , Categories
            self.exts[name] = nt(exten["Name"], exten["Version"], exten["Summary"], exten["Product-Version"])
        
def  Run(args):
    """Execute the STATS PREPROCESS command"""

    args = args[list(args.keys())[0]]

    oobj = Syntax([
        Template("FILTER", subc="",  ktype="str", var="nfilter", islist=False),
        Template("UPDATERPT", subc="", ktype="bool", var="updaterpt"),
        Template("UNINSTALLEDRPT", subc="", ktype="bool", var="uninstalledrpt")
    ])
        
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
