#Author-Ian Shatwell
#Description-Create a basic multi-post setup sheet

import adsk.core, adsk.fusion, adsk.cam, traceback
import os, sys
import time
import pathlib

def WaitForFile(fname):
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

def ParseOnParameter(p):
    onpos = p.find("onParameter(\'")
    if onpos <0:
        return ("","")
    ppair = p[onpos+13:-2]
    splitpos = ppair.find("\',")
    if splitpos <0 :
        return ("","")
    pkey = ppair[:splitpos]
    tvalue = ppair[splitpos+2:].strip()
    if tvalue.find("\'") >= 0:
        pvalue = str(tvalue)
    elif tvalue.find(".") >= 0:
        pvalue = float(tvalue)
    elif str(tvalue).isdigit():
        pvalue = int(tvalue)
    else:
        pvalue = tvalue
    return pkey, pvalue

def run(context):
    ui = None
    allstocksizes = {}
    allstocklimits = {}
    alltools = {}
    setuptools = {}
    alloperations = {}
    allstrategies = {}
    allparameters = {}
    tempfiles = []
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
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
            alloperations[setup.name] = []
            allstrategies[setup.name] = []
            allparameters[setup.name] = {}
            setuptools[setup.name] = []
            operations = setup.allOperations
            for operation in operations:
                alloperations[setup.name].append(operation.name)
                # Get operation information via the Dumper post processor
                # Indirect method as not all information is directly exposed by the API
                # Seems to be the only way of getting stock size information
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
                fdump.close()

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
                    if toolnum != 0 or tooldesc != "":
                        if toolnum not in alltools:
                            alltools[toolnum] = tooldesc
                        if toolnum not in setuptools[setup.name]:
                            setuptools[setup.name].append(toolnum)
                    if toolstrat != "":
                        allstrategies[setup.name].append("{} with #{}".format(toolstrat, toolnum))

        # Report wanted information
        msg = doc.name + "\nStock:\n"
        for stock in allstocksizes:
            msg += "\t"+stock+":\n"
            msg += "\t\tSize: "+allstocksizes[stock]+"\n"
            msg += "\t\t"+allstocklimits[stock]+"\n"
        msg += "\nOperations:\n"
        for setup in allstrategies:
            msg += "\t"+setup+"\n"
            msg += "\t\tTools: "
            for t in setuptools[setup]:
                msg += "#{} ".format(t)
            msg += "\n"
            for n in range(len(allstrategies[setup])):
                msg += "\t\t"+alloperations[setup][n]+": "+allstrategies[setup][n]+"\n"
        msg += "\nFull tool list:\n"
        for tool in alltools:
            msg += "\t#{}: {}\n".format(tool, alltools[tool])

        # Write it to a file as well
        homedir = pathlib.Path.home()
        foutput = open(os.path.join(homedir,doc.name+"_setup.txt"),"w")
        foutput.write(msg)

        # Add other parsed information
        #foutput.write("\n\n\n")
        #for setup in allparameters:
        #    for pkey in allparameters[setup]:
        #        pvalue = allparameters[setup][pkey]
        #        foutput.write("{}|{} = {}\n".format(setup, pkey, pvalue))
        foutput.flush()
        foutput.close()

        # Open the folder containing the file
        if (os.name == 'posix'):
            os.system('open "%s"' % homedir)
        elif (os.name == 'nt'):
            os.startfile(homedir)

        # Display the information on screen
        ui.messageBox(msg, doc.name)

        #Clear up temporary files
        for f in tempfiles:
            os.remove(f)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
