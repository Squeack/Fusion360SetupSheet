#Author-Ian Shatwell
#Description-Create a basic multi-post setup sheet

import adsk.core, adsk.fusion, adsk.cam, traceback
import os, sys, re
import math
import time
import pathlib

THISSCRIPT = "Setup Sheet Generator v2 (c) Ian Shatwell 2020"

# Set these to True or False (case sensitive) to enable or disable output
TXTOUTPUT = False
HTMLOUTPUT = True
SCREENOUTPUT = False

PARAMETER_REGEX = r"\d+:\s*onParameter\(\'([\-\._:\w]+)\',\s*\'?\s*([!-&\(-~\s]*)\'?\)"
LINEAR_REGEX = r"\d+:\s*onLinear\(([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+)\s*\)"
CIRCULAR_REGEX = r"\d+:\s*onCircular\((\w+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+),\s*([-\.0-9eE]+)\s*\)"
POSITION_REGEX = r"\s*STATE\s+position=\[([\-\.0-9]+),\s*([\-\.0-9]+),\s*([\-\.0-9eE]+)\s*\]"

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


def WaitForFile(fname):
    time.sleep(0.2)
    my_file = pathlib.Path(fname)
    # Wait until the file exists
    triesleft = 100
    while triesleft > 0 and not my_file.is_file():
        time.sleep(0.1)
        triesleft -= 1
    if triesleft == 0:
        return False
    # Wait until the file stops growing
    oldsize = my_file.stat().st_size
    time.sleep(0.1)
    newsize = my_file.stat().st_size
    while newsize == 0 or newsize != oldsize:
        time.sleep(0.1)
        oldsize = newsize
        newsize = my_file.stat().st_size
    return True


def ParseOnStatePosition(p):
    try:
        matches = re.finditer(POSITION_REGEX, p, re.ASCII)
        firstmatch=(list(matches))[0]
        x = float(firstmatch.group(1))
        y = float(firstmatch.group(2))
        z = float(firstmatch.group(3))
        return x, y, z
    except:
        # Something went wrong
        adsk.core.Application.get().userInterface.messageBox(p,"Failed to match Position regex")
        return "",""


def ParseOnParameter(p):
    try:
        pkey = ""
        pvalue = ""
        matches = re.finditer(PARAMETER_REGEX, p, re.ASCII)
        firstmatch=(list(matches))[0]
        pkey = firstmatch.group(1)
        pvalue = firstmatch.group(2)
        return pkey,pvalue
    except:
        # Some parameters may have a ' in the text, which will break the regex match
        #adsk.core.Application.get().userInterface.messageBox(p,"Failed to match OnParameter regex")
        return pkey,pvalue


def ParseOnLinear(p, ox, oy, oz, of):
    try:
        matches = re.finditer(LINEAR_REGEX, p, re.ASCII)
        firstmatch=(list(matches))[0]
        x = float(firstmatch.group(1))
        y = float(firstmatch.group(2))
        z = float(firstmatch.group(3))
        f = float(firstmatch.group(4))
        dx = x - ox
        dy = y - oy
        dz = z - oz
        dist = math.sqrt(dx*dx + dy*dy + dz*dz)
        dur = dist / f
        return x, y, z, f, dist, dur
    except:
        # Something went wrong
        adsk.core.Application.get().userInterface.messageBox(p,"Failed to match OnLinear regex")
        raise
        return ox, oy, oz, of, 0, 0


def ParseOnCircular(p, ox, oy, oz, of):
    # Does not correctly calculate angle difference for multiple revolutions, such as a helix ramp
    try:
        matches = re.finditer(CIRCULAR_REGEX, p, re.ASCII)
        firstmatch=(list(matches))[0]
        cw = bool(firstmatch.group(1))
        cx = float(firstmatch.group(2))
        cy = float(firstmatch.group(3))
        cz = float(firstmatch.group(4))
        x = float(firstmatch.group(5))
        y = float(firstmatch.group(6))
        z = float(firstmatch.group(7))
        f = float(firstmatch.group(8))
        # Find start and finish offset from centre
        sdx = ox - cx
        sdy = oy - cy
        sdz = oz - cz
        fdx = x - cx
        fdy = y - cy
        fdz = z - cz
        dx = abs(x - ox)
        dy = abs(y - oy)
        dz = abs(z - oz)
        # Find which plane the rotation is in
        axis = "Z"
        # Look for the smallest axis change and assume a circle or helix around that
        if dx <= dy and dx <= dz: axis = "X"
        if dy <= dx and dy <= dz: axis = "Y"
        if dz <= dx and dz <= dy: axis = "Z"
        if axis == "X":
            r1 = math.sqrt(sdy*sdy + sdz*sdz)
            a1 = math.atan2(sdy, sdz)
            a2 = math.atan2(fdy, fdz)
            dother = dx
        elif axis == "Y":
            r1 = math.sqrt(sdx*sdx + sdz*sdz)
            a1 = math.atan2(sdz, sdx)
            a2 = math.atan2(fdz, fdx)
            dother = dy
        else:
            r1 = math.sqrt(sdx*sdx + sdy*sdy)
            a1 = math.atan2(sdy, sdx)
            a2 = math.atan2(fdy, fdx)
            dother = dz
        # How far around the circle?
        da = diffAngle(cw, a1, a2)
        dist = da * r1
        if dother != 0:
            dist = math.sqrt(dist * dist + dother * dother)
        dur = dist / f
        return x, y, z, f, dist, dur
    except:
        # Something went wrong
        adsk.core.Application.get().userInterface.messageBox(p,"Failed to match OnCircular regex")
        raise
        return ox, oy, oz, of, 0, 0


def run(context):
    ui = None
    allstocksizes = {}
    allstocklimits = {}
    alltools = {}
    setuptools = {}
    operationtools = {}
    alloperations = {}
    allparameters = {}
    tempfiles = []
    distances = {}
    durations = {}
    tooldistances = {}
    tooldurations = {}
    try:
        # Get access to various application and document levels
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
        for setup in cam.setups:
            if not setup.isValid:
                ui.messageBox("Invalid setup", setup.name)
                continue
            alloperations[setup.name] = {}
            allparameters[setup.name] = {}
            setuptools[setup.name] = []
            operationtools[setup.name] = {}
            distances[setup.name] = {}
            durations[setup.name] = {}
            operations = setup.allOperations
            for operation in operations:
                alloperations[setup.name][operation.name] = {}
                distances[setup.name][operation.name]=0
                durations[setup.name][operation.name]=0
                # Get operation information via the Dumper post processor
                # Indirect method as not all information is directly exposed by the API
                # This seems to be the only way of getting stock size information at least
                programName = 'postdump_'+doc.name+"_"+setup.name+"_"+operation.name
                outputFolder = cam.temporaryFolder
                firstSetupOperationType = cam.setups.item(0).operationType
                postConfig = os.path.join(cam.genericPostFolder, 'dump.cps') 
                units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput
                postInput = adsk.cam.PostProcessInput.create(programName, postConfig, outputFolder, units)
                postInput.isOpenInEditor = False
                # create the post properties
                postProperties = adsk.core.NamedValues.create()
                # add the post properties to the post process input
                postInput.postProperties = postProperties
                # Process individual operation
                if operation.hasToolpath == True:
                    cam.postProcess(operation, postInput)
                else:
                    ui.messageBox('Operation {} has no toolpath to post'.format(operation.name))
                    continue

                # Let the file processing finish
                fname = os.path.join(outputFolder,programName+".dmp")
                if not WaitForFile(fname):
                    continue

                # Read the dump back in and look for key information
                tempfiles.append(fname)
                fdump = open(fname,"r")
                lnum = 0
                stockline=""
                toolnum = 0
                tooldesc = ""
                toolstrat = ""
                tooltype = ""
                toolunit = ""
                toolcool = ""
                tooldia = 0
                toollen = 0
                toolflutelen = 0
                toolflutes = 0
                toolshaft = 0
                posx = 0
                posy = 0
                posz = 0
                feedspeed = 1
                minspeed = float('inf')
                maxspeed = 0
                distance = 0
                cuttime = 0

                # Read each line
                for dline in fdump.readlines():
                    lnum += 1
                    if dline.find("onParameter(") >= 0:
                        pkey, pvalue = ParseOnParameter(dline)
                        if pkey != "":
                            # A possibly useful bit of information, so cache it
                            allparameters[setup.name][operation.name+"|"+pkey] = pvalue
                            # Recognise what we need
                            if pkey == "stock": stockline = pvalue
                            if pkey == "operation:tool_number": toolnum = pvalue
                            if pkey == "operation:tool_description": tooldesc = pvalue
                            if pkey == "operation-strategy": toolstrat = pvalue
                            if pkey == "operation:tool_type": tooltype = pvalue
                            if pkey == "operation:tool_diameter": tooldia = pvalue
                            if pkey == "operation:tool_bodyLength": toollen = pvalue
                            if pkey == "operation:tool_fluteLength": toolflutelen = pvalue
                            if pkey == "operation:tool_numberOfFlutes": toolflutes = pvalue
                            if pkey == "operation:tool_shaftDiameter": toolshaft = pvalue
                            if pkey == "operation:tool_unit": toolunit = pvalue
                            if pkey == "operation:tool_coolant": toolcool = pvalue
                    if dline.find("STATE position") >= 0:
                        posx, posy, posz = ParseOnStatePosition(dline)
                    if dline.find("onLinear") >= 0:
                        posx, posy, posz, feedspeed, dist, dur = ParseOnLinear(dline, posx, posy, posz, feedspeed)
                        distance += dist
                        cuttime += dur
                        if feedspeed > maxspeed: maxspeed = feedspeed
                        if feedspeed < minspeed: minspeed = feedspeed
                    if dline.find("onCircular") >= 0:
                        posx, posy, posz, feedspeed, dist, dur = ParseOnCircular(dline, posx, posy, posz, feedspeed)
                        distance += dist
                        cuttime += dur
                        if feedspeed > maxspeed: maxspeed = feedspeed
                        if feedspeed < minspeed: minspeed = feedspeed
                fdump.close()
                operationtools[setup.name][operation.name] = toolnum

                if stockline == "":
                    ui.messageBox(str(lnum)+" lines searched, but no stock information found")
                else:
                    # Rearrange the stock information to a useful format
                    stocksize = stockline.replace("\'","").replace("(","").replace(")","")
                    msg = "Stock is {}\n".format(stocksize)
                    stocksize = stocksize.split(",")
                    minx = stocksize[0]
                    miny = stocksize[1]
                    minz = stocksize[2]
                    maxx = stocksize[3]
                    maxy = stocksize[4]
                    maxz = stocksize[5]
                    allstocksizes[setup.name]="{} x {} x {}".format(float(maxx)-float(minx), float(maxy)-float(miny), float(maxz)-float(minz))
                    allstocklimits[setup.name]="Lower: {}, {}, {}; Upper: {}, {}, {}".format(minx, miny, minz, maxx, maxy, maxz)

                # Store tooling information related to the operation
                msg += "Tool number: {}\n".format(toolnum)
                msg += "{}".format(tooldesc)
                if toolnum != 0:
                    if toolnum not in alltools:
                        alltools[toolnum] = {}
                        alltools[toolnum]["description"] = tooldesc
                        alltools[toolnum]["type"] = tooltype
                        alltools[toolnum]["cuttingdiameter"] = tooldia
                        alltools[toolnum]["length"] = toollen
                        alltools[toolnum]["flutelength"] = toolflutelen
                        alltools[toolnum]["numflutes"] = toolflutes
                        alltools[toolnum]["shaftdiameter"] = toolshaft
                        alltools[toolnum]["units"] = toolunit
                        alltools[toolnum]["minspeed"] = minspeed
                        alltools[toolnum]["maxspeed"] = maxspeed
                    if toolnum not in setuptools[setup.name]:
                        setuptools[setup.name].append(toolnum)
                    if toolnum not in tooldistances:
                        tooldistances[toolnum] = 0
                    if toolnum not in tooldurations:
                        tooldurations[toolnum] = 0
                alloperations[setup.name][operation.name]["tool"] = toolnum
                alloperations[setup.name][operation.name]["strategy"] = toolstrat
                alloperations[setup.name][operation.name]["minspeed"] = minspeed
                alloperations[setup.name][operation.name]["maxspeed"] = maxspeed
                alloperations[setup.name][operation.name]["coolant"] = toolcool
                distances[setup.name][operation.name]=distance
                durations[setup.name][operation.name]=cuttime
                tooldistances[toolnum] = tooldistances[toolnum] + distance
                tooldurations[toolnum] = tooldurations[toolnum] + cuttime

        # Report wanted information
        msg = doc.name + "\nStock:\n"
        for stock in allstocksizes:
            msg += "\t"+stock+":\n"
            msg += "\t\tSize: "+allstocksizes[stock]+"\n"
            msg += "\t\t"+allstocklimits[stock]+"\n"
        msg += "\nOperations:\n"
        for setup in alloperations:
            d = 0
            t = 0
            for op in distances[setup]:
                d += distances[setup][op]
            for op in durations[setup]:
                t += durations[setup][op]
            msg += "\t{} ({:.0f}mm in {:d}m{:d}s not allowing for acc/deceleration)\n".format(setup, d, int(t), int(t*60) % 60)
            msg += "\t\tTools: "
            for t in setuptools[setup]:
                msg += "#{} ".format(t)
            msg += "\n"
            operationsdetails = alloperations[setup]
            for op in operationsdetails:
                msg += "\t\t{}: {} with #{}\n".format(op, operationsdetails[op]["strategy"], operationsdetails[op]["tool"])
            
        msg += "\nFull tool list:\n"
        for tool in alltools:
            msg += "\t#{}: {}\n".format(tool, alltools[tool]["description"])

        if TXTOUTPUT:
            # Write it to a file as well
            homedir = pathlib.Path.home()
            outputname = os.path.join(homedir,doc.name+"_setup.txt")
            foutput = open(outputname,"w")
            foutput.write(msg)

            # Add other parsed information
            foutput.write("\n\n\n")
            for setup in allparameters:
                for pkey in allparameters[setup]:
                    pvalue = allparameters[setup][pkey]
                    foutput.write("{}|{} = {}\n".format(setup, pkey, pvalue))
            foutput.write("\nProduced by {}\n".format(THISSCRIPT))
            foutput.flush()
            foutput.close()
            # Open the file
            OpenFile(outputname)

        if HTMLOUTPUT:
            # Write HTML version
            homedir = pathlib.Path.home()
            outputname = os.path.join(homedir,doc.name+" setupsheet.html")
            foutput = open(outputname,"w")
            title = "Setup Sheet for {}".format(doc.name)
            HTMLheader(foutput,title)
            foutput.write("<body>\n")
            HTMLBodyTitle(foutput,title)
            HTMLsetups(foutput, allstocksizes, allstocklimits)
            foutput.write("<br><br>\n")
            HTMLtools(foutput, alltools, tooldistances, tooldurations)
            foutput.write("<br><br>\n")
            HTMLoperations(foutput, alltools, alloperations, allparameters, distances, durations)
            foutput.write("\n<br><div align=\"left\" style=\"font-size:5pt; color: PowderBlue\">Produced by {}</div>\n".format(THISSCRIPT))
            foutput.write("</body>")
            foutput.flush()
            foutput.close()
            # Open the file
            OpenFile(outputname)

        if SCREENOUTPUT:
            # Display the information on screen
            ui.messageBox(msg, doc.name)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

    #Clear up temporary files
    sleep(1)
    for f in tempfiles:
        os.remove(f)


def HTMLheader(f, title):
    f.write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\n")
    f.write("                      \"http://www.w3.org/TR/1999/REC-html401-19991224/loose.dtd\">\n")
    f.write("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n")
    f.write(STYLESHEET)
    f.write("<title>{}</title>\n".format(title))
    f.write("</head>\n")

def HTMLBodyTitle(f, title):
    f.write("<h1>{}</h1>\n".format(title))

def HTMLsetups(f, sizes, limits):
    f.write("<table class=\"setup\" cellspacing=0 align=\"center\">\n")
    f.write("<tr><th colspan=3>Stocks</th></tr>\n")
    f.write("<tr class=\"lined\"><td class=\"description\">Setup name</td><td class=\"description\">Stock Size</td><td class=\"description\">Limits</td></tr>")
    for setup in sizes:
        f.write("<tr><td>")
        f.write("<div class=\"value\"><b>{}</b></div>".format(setup))
        f.write("</td>\n<td>")
        f.write("<div class=\"value\">{}</div>".format(sizes[setup]))
        f.write("</td>\n<td>")
        f.write("<div class=\"value\">")
        limitlist = limits[setup].split(";")
        for n in range(len(limitlist)):
            limit = limitlist[n].strip()
            f.write("{}".format(limit))
            if n < len(limitlist)-1 :
                f.write("<br>")
        f.write("</div></td></tr>\n")
    f.write("</table>\n")

def HTMLtools(f, toollist, tooldist, tooltime):
    f.write("<table class=\"setup\" cellspacing=0 align=\"center\">\n")
    f.write("<tr><th colspan=2>Tools</th></tr>\n")
    for t in toollist:
        if toollist[t]["units"] == "millimeters" or toollist[t]["units"] == "mm":
            units = "mm"
        else:
            units = "in"
        f.write("<tr class=\"tool\"><td colspan=2 align=\"left\"><b>#{}</b></td></tr>\n".format(t))
        f.write("<tr class=\"lined\"><td class=\"description\">Description</td><td class=\"description\">Usage</td></tr>")
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
        f.write("<div class=\"value\">{}{}</div>".format(toollist[t]["length"], units))
        f.write("<br>")
        #f.write("<div class=\"description\">Flute length: </div>")
        #f.write("<div class=\"value\">{}{}</div>".format(toollist[t]["flutelength"], units))
        #f.write("<br>")
        f.write("<div class=\"description\">Number of flutes: </div>")
        f.write("<div class=\"value\">{}</div>".format(toollist[t]["numflutes"]))
        f.write("<br>")
        f.write("<div class=\"description\">Shaft diameter: </div>")
        f.write("<div class=\"value\">{}{}</div>".format(toollist[t]["shaftdiameter"], units))
        f.write("</td><td>")
        f.write("<div class=\"description\">Travel: </div>")
        f.write("<div class=\"value\">{:.0f}{}</div>".format(tooldist[t], units))
        f.write("<br>")
        f.write("<div class=\"description\">Time: </div>")
        f.write("<div class=\"value\">{}m{}s</div>".format(int(tooltime[t]), int(tooltime[t]*60) % 60))
        f.write("<br>")
        f.write("<div class=\"description\">Minimum cutting speed: </div>")
        f.write("<div class=\"value\">{:.1f}{}/min</div>".format(toollist[t]["minspeed"], units))
        f.write("<br>")
        f.write("<div class=\"description\">Maximum cutting speed: </div>")
        f.write("<div class=\"value\">{:.1f}{}/min</div>".format(toollist[t]["maxspeed"], units))
        f.write("</td></tr>")
    f.write("</table>\n")

def HTMLoperations(f, toollist, operations, allparams, dist, dur):
    for setup in operations:
        f.write("<table class=\"setup\" cellspacing=0 align=\"center\">\n")
        f.write("<tr><th colspan=3>Operations for {}</th></tr>\n".format(setup))
        setupparams = allparams[setup]
        n = 0 
        for op in operations[setup]:
            n += 1
            basename="{}|".format(op)
            f.write("<tr class=\"tool\"><td colspan=2 align=\"left\">Operations {}/{}: {}</td>".format(n, len(operations[setup]), op))
            toolnum = operations[setup][op]["tool"]
            units = toollist[toolnum]["units"]
            if units == "millimeters" or units == "mm":
                units = "mm"
            else:
                units = "in"
            f.write("<td>#{}</td></tr>\n".format(toolnum))
            f.write("<tr><td>")
            f.write("<div class=\"description\">Strategy: </div><div class=\"value\">{}</div>".format(operations[setup][op]["strategy"]))
            if basename+"operation:tolerance" in setupparams:
                f.write("<br><div class=\"description\">Tolerance: </div><div class=\"value\">{}{}</div>".format(setupparams[basename+"operation:tolerance"],units))
            if basename+"operation:maximumStepdown" in setupparams:
                f.write("<br><div class=\"description\">Max stepdown: </div><div class=\"value\">{}{}</div>".format(setupparams[basename+"operation:maximumStepdown"],units))
            if basename+"operation:maximumStepover" in setupparams:
                f.write("<br><div class=\"description\">Max stepover: </div><div class=\"value\">{}{}</div>".format(setupparams[basename+"operation:maximumStepover"],units))
            if basename+"operation:optimalLoad" in setupparams:
                f.write("<br><div class=\"description\">Optical load: </div><div class=\"value\">{}{}</div>".format(setupparams[basename+"operation:optimalLoad"],units))
            f.write("</td><td>")
            f.write("<div class=\"description\">Distance: </div><div class=\"value\">{:.0f}{}</div>".format(dist[setup][op],units))
            duration = dur[setup][op]
            f.write("<br><div class=\"description\">Time: </div><div class=\"value\">{}m{}s</div>".format(int(duration), int(duration*60) % 60))
            f.write("<br><div class=\"description\">Minimum speed: </div><div class=\"value\">{:.1f}{}/min</div>".format(operations[setup][op]["minspeed"],units))
            f.write("<br><div class=\"description\">Maximum speed: </div><div class=\"value\">{:.1f}{}/min</div>".format(operations[setup][op]["maxspeed"],units))
            f.write("<br><div class=\"description\">Coolant: </div><div class=\"value\">{}</div>".format(operations[setup][op]["coolant"]))
            f.write("</td><td>")
            f.write("<div class=\"description\">Type: </div>")
            f.write("<div class=\"value\">{}</div>".format(toollist[toolnum]["type"]))
            f.write("<br>")
            f.write("<div class=\"description\">Diameter: </div>")
            f.write("<div class=\"value\">{}{}</div>".format(toollist[toolnum]["cuttingdiameter"], units))
            f.write("<br>")
            f.write("<div class=\"description\">External length: </div>")
            f.write("<div class=\"value\">{}{}</div>".format(toollist[toolnum]["length"], units))
            f.write("<br>")
            f.write("<div class=\"description\">Number of flutes: </div>")
            f.write("<div class=\"value\">{}</div>".format(toollist[toolnum]["numflutes"]))
            f.write("<br>")
            f.write("<div class=\"value\">{}</div>".format(toollist[toolnum]["description"]))
            f.write("</td></tr>")
        f.write("</table><br>")
        