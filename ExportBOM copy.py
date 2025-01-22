#Author-Autodesk Inc.
#Description-Extract BOM information from active design.

import adsk.core, adsk.fusion, traceback, time, csv, os
# import pandas as pd # type: ignore

# from openpyxl import Workbook # type: ignore
# from openpyxl.drawing.image import Image # type: ignore

defaultExportToCSVEnabbled = False
def spacePadRight(value, length):
    pad = ''
    if type(value) is str:
        paddingLength = length - len(value) + 1
    else:
        paddingLength = length - value + 1
    while paddingLength > 0:
        pad += ' '
        paddingLength -= 1

    return str(value) + pad

def walkThrough(bom):
    mStr = ''
    for item in bom:
        mStr += spacePadRight(item['name'], 25) + str(spacePadRight(item['instances'], 15))+ str(spacePadRight(item['mat'], 15)) + str(item['volume']) + '\n'
    return mStr

def buildCSV1(bom, fileName):
    with open(fileName + '.csv', 'w', encoding='utf-8', newline='') as csvfile:
        writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, delimiter=';')
        writer.writerow(['name', 'instances','mat', 'volume'])
        for item in bom:
            name =  item['name']
            writer.writerow([name, item['instances'],item['mat'], item['volume']])
def buildCSV2(bom, fileName):
    # Use a comma as the delimiter for proper Excel compatibility
    with open(fileName + '.csv', 'w', encoding='utf-8', newline='') as csvfile:
        writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, delimiter=',')
        # Write the header row
        writer.writerow(['name', 'instances', 'mat', 'volume'])
        # Write each item in BOM
        for item in bom:
            writer.writerow([item['name'], item['instances'], item['mat'], item['volume']])

def buildCSV(bom, imageDirectory, fileName):
    with open(fileName + '.csv', 'w', encoding='utf-8', newline='') as csvfile:
        writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, delimiter=',')
        # Add a new column for the image path
        writer.writerow(['name', 'instances', 'mat', 'volume', 'image_path'])
        
        for item in bom:
            name = item['name']
            # Construct the image path
            image_path = os.path.join(imageDirectory, f"{item['component'].id}.png")
            writer.writerow([name, item['instances'], item['mat'], item['volume'], image_path])

def buildHTMLWithImages(app,bom, imageDirectory, fileName):
    # path, dir = os.path.split(imageDirectory)
    with open(fileName + '.html', 'w', encoding='utf-8') as html_file:
        html_file.write('<html><body>\n')
        html_file.write('<table border="1">\n')
        html_file.write('<tr><th>Name</th><th>Instances</th><th>Material</th><th>Volume</th><th>Image</th></tr>\n')
        
        for item in bom:
            image_path = f"{imageDirectory}/{item['component'].id}.png"
            app.log(image_path)
            html_file.write('<tr>')
            html_file.write(f'<td>{item["name"]}</td>')
            html_file.write(f'<td>{item["instances"]}</td>')
            html_file.write(f'<td>{item["mat"]}</td>')
            html_file.write(f'<td>{item["volume"]}</td>')
            
            if os.path.exists(image_path):
                html_file.write(f'<td><img src="{image_path}" alt="Image" width="50" height="50"></td>')
            else:
                html_file.write('<td>Image not found</td>')
            
            html_file.write('</tr>\n')
        
        html_file.write('</table>\n')
        html_file.write('</body></html>\n')

def setGridDisplay(turnOn):
    app = adsk.core.Application.get()
    ui  = app.userInterface

    cmdDef = ui.commandDefinitions.itemById('ViewLayoutGridCommand')
    listCntrlDef = adsk.core.ListControlDefinition.cast(cmdDef.controlDefinition)
    layoutGridItem = listCntrlDef.listItems.item(0)
    
    if turnOn:
        layoutGridItem.isSelected = True
    else:
        layoutGridItem.isSelected = False 

def takeImage(app,ui,component, occs, path):

    #HideAll(occs)

    cameraTarget = False
    occurrence = False

    for occ in occs:
        comp = occ.component

        if comp == component and not cameraTarget:
            cameraTarget = adsk.core.Point3D.create(occ.transform.translation.x, occ.transform.translation.y, occ.transform.translation.z)
            #comp.isBodiesFolderLightBulbOn = True
            occurrence = occ
        
            

    if cameraTarget:

        if occurrence.assemblyContext and occurrence.assemblyContext.isReferencedComponent:
            occurrence.assemblyContext.isIsolated = True
            
            hiddenChildren = []
            for childOccurrence in occurrence.assemblyContext.component.allOccurrences:
                hiddenChild = childOccurrence.component

                if hiddenChild != component and hiddenChild.isBodiesFolderLightBulbOn:
                    hiddenChildren.append(hiddenChild)
                    hiddenChild.isBodiesFolderLightBulbOn = False

        else:
            occurrence.isIsolated = True

        setGridDisplay(False)

        viewport = app.activeViewport
        camera = viewport.camera

        camera.target = cameraTarget
        camera.isFitView = True
        camera.isSmoothTransition = False
        # camera.eye = cameraBackup.eye
        camera.eye = adsk.core.Point3D.create(100 + cameraTarget.x, -100 + cameraTarget.y, 100 + cameraTarget.z)

        app.activeViewport.camera = camera
        
        app.activeViewport.refresh()
        adsk.doEvents()

        success = app.activeViewport.saveAsImageFile(path + '/' + component.id  + '.png', 128, 128)
        if not success:
            ui.messageBox('Failed saving viewport image.')
        
        if occurrence.assemblyContext and occurrence.assemblyContext.isReferencedComponent:

            occurrence.assemblyContext.isIsolated = False  

            for hiddenChild in hiddenChildren:
                hiddenChild.isBodiesFolderLightBulbOn = True            

        else:
            occurrence.isIsolated = False

def Unisolate(occs):
    for occ in occs:
        occ.isLightBulbOn = True

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        title = 'Extract BOM'
        if not design:
            ui.messageBox('The DESIGN workspace must be active when running this script.', title)
            return

        # Get all occurrences in the root component of the active design
        root = design.rootComponent
        occs = root.allOccurrences

        visibleTopLevelComp = []
        for occ in root.occurrences:
            if occ.isLightBulbOn:
                visibleTopLevelComp.append(occ)

        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = " filename"
        fileDialog.filter = 'html (*.html)'
        fileDialog.initialFilename =  product.rootComponent.name
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showSave()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
            path, file = os.path.split(filename)
            #path = os.path.dirname(filename)
            dst_directory = os.path.splitext(filename)[0] + '_files'
        
        # Gather information about each unique component
        bom = []
        for occ in occs:
            comp = occ.component
            jj = 0
            for bomI in bom:
                if bomI['component'] == comp:
                    # Increment the instance count of the existing row.
                    bomI['instances'] += 1
                    break
                jj += 1

            if jj == len(bom):
                # Gather any BOM worthy values from the component
                volume = 0
                bodies = comp.bRepBodies
                for bodyK in bodies:
                    if bodyK.isSolid:
                        volume += bodyK.volume
                
                # Add this component to the BOM
                bom.append({
                    'component': comp,
                    'name': comp.name,
                    'instances': 1,
                    'mat':comp.material.name,
                    'volume': volume
                    
                })

        for bomItem in bom:
            takeImage(app,ui,bomItem['component'], occs, dst_directory)
            Unisolate(visibleTopLevelComp)
        # ShowAll(occs)

        # Display the BOM
        title = spacePadRight('Name', 25) + spacePadRight('Instances', 15)+ str(spacePadRight('Material', 15)) + 'Volume\n'
        msg = title + '\n' + walkThrough(bom)
        buildCSV(bom,dst_directory, os.path.splitext(filename)[0] + '_bom')
        buildHTMLWithImages(app,bom,dst_directory, os.path.splitext(filename)[0] + '_html')
        # app.log(walkThrough(bom))
        app.log(os.path.splitext(filename)[0] + '_bom')
        ui.messageBox(msg, 'Bill Of Materials')

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
