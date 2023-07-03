#Author-Ian Shatwell
#Description-Create a basic multi-post setup sheet

import adsk.core, adsk.fusion, adsk.cam, traceback
import os
import math
import pathlib

THISSCRIPT = "Setup Sheet Generator v3 (c) Ian Shatwell 2023"

# Set these to True or False (case sensitive) to enable or disable output
TXTOUTPUT = False
HTMLOUTPUT = True
SCREENOUTPUT = False

STYLESHEET = """
    <style type="text/css">
    body {background-color:white; font-family: Arial, Helvetica, sans-serif;}
    h1 {font-size: 16pt;text-align: center;}
    table { border: none; border-spacing: 0;}
    table.setup, table.sheet {width: 18cm; border: 1px solid Black;}
    table.info {padding-top: 0.1cm;}
    table.info td { padding-left: 0.1cm;}
    tr {border: 1px solid Black; page-break-inside: avoid; padding-top: 30px; padding-bottom: 20px; white-space: nowrap;}
    tr.lined td {border-bottom: 1px solid Gray}
    tr.tool td {background-color: #e0e0f0; border-bottom: 1px solid Gray; border-top: 1px solid Gray;}
    th {background-color: #d0d0f0; border-bottom: 1px solid Gray; border-top: 1px solid Gray;}
    td {font-size: 9pt; vertical-align: top;}
    td .description {display: inline; font-variant: small-caps;}
    td .value {display: inline; font-family: Geneva, sans-serif; color: #404060;}
    </style>\n
"""


def floatMatch(f1,f2,e=0.00001):
    return abs(f1-f2) < e


def diffAngle(cw, a1, a2):
    # Angles in range -pi to pi. 0 = across, +ve =up
    # Return angle from a1 to a2 in direction specified
    pi2 = 2 * math.pi
    if cw:
        while(a1 < a2):
            a1 += pi2
        da = a1 - a2
    else:
        while(a2 < a1):
            a2 += pi2
        da = a2 - a1
    while (da > pi2):
        da -= pi2
    while (da <= 0):
        da += pi2
    return da


def OpenFile(fname):
    if (os.name == 'posix'):
        os.system('open "%s"' % fname)
    elif (os.name == 'nt'):
        os.startfile(fname)

def MinimumOf(v1,v2):
    if v1 is None:
        retval = v2
    else:
        retval = v1
        if v2 is not None:
            if v2 < v1:
                retval = v2
    return retval


def MaximumOf(v1,v2):
    if v1 is None:
        retval = v2
    else:
        retval = v1
        if v2 is not None:
            if v2 > v1:
                retval = v2
    return retval
    

def FindParameterValue(params, pname):
    p = params.itemByName(pname)
    if not p.isValid:
        return None
    exp = p.expression
    v=p.value.value
    # Does it copy another parameter value?
    deeper = params.itemByName(exp)
    if deeper is not None:
        return FindParameterValue(params, deeper.name)
    return v


def HTMLheader(f, title):
    f.write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\n")
    f.write("                      \"http://www.w3.org/TR/1999/REC-html401-19991224/loose.dtd\">\n")
    f.write("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n")
    f.write(STYLESHEET)
    f.write("<title>{}</title>\n".format(title))
    f.write("</head>\n")


def HTMLBodyTitle(f, title):
    f.write("<h1>{}</h1>\n".format(title))


def HTMLsetups(f, allsetups, allstocksizes):
    f.write("<table class=\"setup\" cellspacing=0 align=\"center\">\n")
    f.write("<tr><th colspan=3>Stocks</th></tr>\n")
    f.write("<tr class=\"lined\"><td class=\"description\">Setup name</td><td class=\"description\">Stock Size</td><td class=\"description\">Limits</td></tr>")
    for setupnum in allsetups:
        minx, maxx, miny, maxy, minz, maxz = allstocksizes[setupnum]
        f.write("<tr><td>")
        f.write("<div class=\"value\"><b>{}</b></div>".format(allsetups[setupnum]))
        f.write("</td>\n<td>")
        f.write("<div class=\"value\">{:.1f} x {:.1f} x {:.1f}</div>".format(maxx-minx, maxy-miny, maxz-minz))
        f.write("</td>\n<td>")
        f.write("<div class=\"value\">")
        f.write("Lower: {:.1f}, {:.1f}, {:.1f}<br>".format(minx, miny, minz))
        f.write("Upper: {:.1f}, {:.1f}, {:.1f}".format(maxx, maxy, maxz))
        f.write("</div></td></tr>\n")
    f.write("</table>\n")


def HTMLtools(f, toollist):
    f.write("<table class=\"setup\" cellspacing=0 align=\"center\">\n")
    f.write("<tr><th>Tools</th></tr>\n")
    for t in toollist:
        print("Output for {}".format(t))
        if toollist[t]["units"] == "millimeters" or toollist[t]["units"] == "mm":
            units = "mm"
        else:
            units = "in"
        f.write("<tr class=\"tool\"><td align=\"left\"><b>#{}</b></td></tr>\n".format(toollist[t]["number"]))
        f.write("<tr><td>")
        f.write("<div class=\"value\"><b>{}</b></div>".format(toollist[t]["description"]))
        f.write("<br>")
        f.write("<div class=\"description\">Type: </div>")
        f.write("<div class=\"value\">{}</div>".format(toollist[t]["type"]))
        f.write("<br>")
        f.write("<div class=\"description\">Cutting diameter: </div>")
        f.write("<div class=\"value\">{}{}</div>".format(toollist[t]["cuttingdiameter"], units))
        f.write("<br>")
        f.write("<div class=\"description\">External length: </div>")
        f.write("<div class=\"value\">{:.1f}{}</div>".format(toollist[t]["externallength"], units))
        f.write("<br>")
        f.write("<div class=\"description\">Number of flutes: </div>")
        f.write("<div class=\"value\">{}</div>".format(toollist[t]["flutes"]))
        f.write("<br>")
        f.write("<div class=\"description\">Shaft diameter: </div>")
        f.write("<div class=\"value\">{}{}</div>".format(toollist[t]["shaftdiameter"], units))
        f.write("</td></tr>")
    f.write("</table>\n")


def HTMLoperations(f, allsetups, setupops):
    for setupnum in allsetups:
        f.write("<table class=\"setup\" cellspacing=0 align=\"center\">\n")
        f.write("<tr><th colspan=3>Operations for {}</th></tr>\n".format(allsetups[setupnum]))
        operations = setupops[setupnum]
        for opnum in operations:
            opname = operations[opnum]["name"]
            opstrategy = operations[opnum]["strategy"]
            optool = operations[opnum]["tool"]
            optolerance = operations[opnum]["tolerance"]
            f.write("<tr class=\"tool\"><td colspan=2 align=\"left\">Operations {}/{}: {}</td>".format(opnum+1, len(operations), opname))
            units = optool["units"]
            if units == "millimeters" or units == "mm" or units == "'millimeters'" or units == "'mm'":
                units = "mm"
            else:
                units = "in"
            f.write("<td>#{}</td></tr>\n".format(optool["number"]))
            f.write("<tr><td>")
            f.write("<div class=\"description\">Strategy: </div><div class=\"value\">{}</div>".format(opstrategy))
            f.write("<br><div class=\"description\">Tolerance: </div><div class=\"value\">{}{}</div>".format(optolerance,units))
            f.write("</td><td>")
            f.write("<div class=\"description\">Cutting distance: </div><div class=\"value\">{:.0f}{}</div>".format(optool["feedD"],units))
            duration = optool["totalT"] / 60.0
            f.write("<br><div class=\"description\">Total time: </div><div class=\"value\">{}m{}s</div>".format(int(duration), int(duration*60) % 60))
            f.write("<br><div class=\"description\">Coolant: </div><div class=\"value\">{}</div>".format(optool["coolant"]))
            f.write("</td><td>")
            f.write("<div class=\"description\">Type: </div>")
            f.write("<div class=\"value\">{}</div>".format(optool["type"]))
            f.write("<br>")
            f.write("<div class=\"description\">Diameter: </div>")
            f.write("<div class=\"value\">{}{}</div>".format(optool["cuttingdiameter"], units))
            f.write("<br>")
            f.write("<div class=\"description\">External length: </div>")
            f.write("<div class=\"value\">{}{}</div>".format(optool["externallength"], units))
            f.write("<br>")
            f.write("<div class=\"description\">Flutes: </div>")
            f.write("<div class=\"value\">{}</div>".format(optool["flutes"]))
            f.write("<br>")
            f.write("<div class=\"value\">{}</div>".format(optool["description"]))
            f.write("</td></tr>")
        f.write("</table><br>")
        

def NewToolEntry():
    entry={}
    entry["number"] = 0
    entry["description"] = ""
    entry["type"] = ""
    entry["cuttingdiameter"] = 0.0
    entry["shaftdiameter"] = 0.0
    entry["externallength"] = 0.0
    entry["flutes"] = 0
    entry["coolant"] = ""
    entry["units"] = "millimeters"
    return entry


def ToolFromParams(params):
    entry = NewToolEntry()
    entry["number"] = int(FindParameterValue(params, "tool_number"))
    entry["description"] = FindParameterValue(params, "tool_description")
    entry["type"] = FindParameterValue(params, "tool_type")
    entry["cuttingdiameter"] = round(10.0 * float(FindParameterValue(params, "tool_diameter")),6)
    entry["shaftdiameter"] = round(10.0 * float(FindParameterValue(params, "tool_shaftDiameter")),6)
    entry["externallength"] = round(10.0 * float(FindParameterValue(params, "tool_bodyLength")),6)
    entry["flutes"] = int(FindParameterValue(params, "tool_numberOfFlutes"))
    entry["coolant"] = FindParameterValue(params, "tool_coolant")
    entry["units"] = FindParameterValue(params, "tool_unit")
    return entry


def run(context):
    ui = None
    allsetups = {}
    allstocksizes = {}
    setupops = {}
    alltools = {}
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        if not TXTOUTPUT and not HTMLOUTPUT and not SCREENOUTPUT:
            ui.messageBox("No output modes are enabled\nWhy are we doing this?", "Waste of time")
            return
        doc = app.activeDocument
        products = doc.products
        product = products.itemByProductType('CAMProductType')
        # check if the document has a CAMProductType
        if product == None:
            ui.messageBox('There are no CAM operations in the active document.  This script requires the active document to contain at least one CAM operation.',
                            'No CAM Operations Exist',
                            adsk.core.MessageBoxButtonTypes.OKButtonType,
                            adsk.core.MessageBoxIconTypes.CriticalIconType)
            return

        cam = adsk.cam.CAM.cast(product)
        # Initial run through looking at stock sizes and what tools are used
        for setupnum in range(len(cam.setups)):
            setup = cam.setups[setupnum]
            allsetups[setupnum] = setup.name
            setupops[setupnum] = {}
            if not setup.isValid:
                ui.messageBox("Invalid setup", setup.name)
                continue
            print("Setup: " + setup.name)
            alloperations = setup.allOperations
            stockminx = None
            stockmaxx = None
            stockminy = None
            stockmaxy = None
            stockminz = None
            stockmaxz = None
            for opnum in range(len(alloperations)):
                op = alloperations[opnum]
                print("Operation: " + op.name)
                if op.isSuppressed:
                    print("  Suppressed")
                    continue
                if not op.isValid or not op.isToolpathValid:
                    print("  Invalid")
                    ui.messageBox("Invalid operation for {}:{}".format(setup.name, op.name), "Operation Error")
                    continue
                if op.hasWarning or op.isGenerating or not op.hasToolpath:
                    print("  Needs to be generated successfully")
                    ui.messageBox("No toolpath for {}:{}".format(setup.name, op.name), "Operation Error")
                    continue
                print("Strategy: " + op.strategy)
                # attrs = op.attributes
                # for attr in attrs:
                #     print("  Attribute: {}={}".format(attr.name, attr.value))
                params = op.parameters
                setupops[setupnum][opnum] = {}
                setupops[setupnum][opnum]["name"] = op.name
                setupops[setupnum][opnum]["strategy"] = op.strategy
                setupops[setupnum][opnum]["tolerance"] = round(FindParameterValue(params, "tolerance"),6)
                stockminx = MinimumOf(stockminx, float(params.itemByName("stockXLow").expression))
                stockmaxx = MaximumOf(stockmaxx, float(params.itemByName("stockXHigh").expression))
                stockminy = MinimumOf(stockminy, float(params.itemByName("stockYLow").expression))
                stockmaxy = MaximumOf(stockmaxy, float(params.itemByName("stockYHigh").expression))
                stockminz = MinimumOf(stockminz, float(params.itemByName("stockZLow").expression))
                stockmaxz = MaximumOf(stockmaxz, float(params.itemByName("stockZHigh").expression))
                # print("Stock {} to {}, {} to {}, {} to {}".format(stockminx, stockmaxx, stockminy, stockmaxy, stockminz, stockmaxz))
                allstocksizes[setupnum] = (stockminx, stockmaxx, stockminy, stockmaxy, stockminz, stockmaxz)
                # for param in params:
                    # if param.name[:5] == "tool_":
                    # print("{} = {}".format(param.name, FindParameterValue(params, param.name)))
                optool = ToolFromParams(params)
                cuttingSpeed = FindParameterValue(params, "tool_feedCutting")
                units = optool["units"]
                if units == "millimeters" or units == "mm" or units == "'millimeters'" or units == "'mm'":
                    cuttingSpeed /= 10.0
                else:
                    cuttingSpeed /= 2.54
                mtime = cam.getMachiningTime(op, 100.0, cuttingSpeed,0.0)
                optool["feedD"] = mtime.feedDistance
                optool["feedT"] = mtime.totalFeedTime
                optool["rapidD"] = mtime.rapidDistance
                optool["totalT"] = mtime.machiningTime
                toolID = "{}¦{}¦{}¦{}".format(optool["description"], optool["cuttingdiameter"], optool["type"], optool["number"])
                if toolID not in alltools:
                    alltools[toolID] = optool
                print(toolID)
                print(optool)
                setupops[setupnum][opnum]["tool"] = optool

        print(allstocksizes)
        if HTMLOUTPUT:
            # Write HTML version
            homedir = pathlib.Path.home()
            outputname = os.path.join(homedir,doc.name+" setupsheet.html")
            foutput = open(outputname,"w")
            title = "Setup Sheet for {}".format(doc.name)
            HTMLheader(foutput,title)
            foutput.write("<body>\n")
            HTMLBodyTitle(foutput,title)
            HTMLsetups(foutput, allsetups, allstocksizes)
            foutput.write("<br><br>\n")
            HTMLtools(foutput, alltools)
            foutput.write("<br><br>\n")
            HTMLoperations(foutput, allsetups, setupops)
            foutput.write("\n<br><div align=\"left\" style=\"font-size:5pt; color: PowderBlue\">Produced by {}</div>\n".format(THISSCRIPT))
            foutput.write("</body>")
            foutput.flush()
            foutput.close()
            # Open the file
            OpenFile(outputname)        

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
