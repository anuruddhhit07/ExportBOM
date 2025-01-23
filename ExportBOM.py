import adsk.core, adsk.fusion, traceback, time, csv, os, base64,shutil
from .BOMExporterClass import BOMExporter
# Global variable to control whether volume is included
includeVolume = False  # Set this to False if you do not want to include volume in the BOM

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
        exporter = BOMExporter(includeVolume)
        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = "Save BOM As"
        fileDialog.filter = 'HTML (*.html)'
        fileDialog.initialFilename = product.rootComponent.name
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showSave()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
            
            
            exporter.delete_related_files(filename)
            path, file = os.path.split(filename)
            dst_directory = os.path.splitext(filename)[0] + '_files'

        # Gather information about each unique component
        bom = []
        for occ in occs:
            comp = occ.component
            jj = 0
            for bomI in bom:
                if bomI['component'] == comp:
                    bomI['instances'] += 1
                    break
                jj += 1

            if jj == len(bom):
                volume = 0
                bodies = comp.bRepBodies
                for bodyK in bodies:
                    if bodyK.isSolid:
                        volume += bodyK.volume
                
                bom.append({
                    'component': comp,
                    'name': comp.name,
                    'instances': 1,
                    'mat': comp.material.name,
                    'volume': volume if includeVolume else None
                })

        for bomItem in bom:
            exporter.take_image(app, ui, bomItem['component'], occs, dst_directory)
            exporter.Unisolate(visibleTopLevelComp)

        # Display BOM data and save
        # msg = spacePadRight('Name', 25) + spacePadRight('Instances', 15) + spacePadRight('Material', 15) + 'Volume\n' + walkThrough(bom)
        # Display BOM data and save files
        msg = exporter.space_pad_right('Name', 25) + \
                  exporter.space_pad_right('Instances', 15) + \
                  exporter.space_pad_right('Material', 15) + \
                  ('Volume\n' if includeVolume else '\n') + \
                  exporter.walk_through(bom)
        
        exporter.build_csv(bom, dst_directory, os.path.splitext(filename)[0] + '_bom')
        exporter.build_html_with_images(app, bom, dst_directory, os.path.splitext(filename)[0] + '_html', editable=False)
        exporter.buildHTMLWithImagesEditableCSV(app, bom, dst_directory, os.path.splitext(filename)[0] + '_html', editable=True)
        ui.messageBox(msg, 'Bill Of Materials')

    except Exception as e:
        if ui:
            ui.messageBox(f'Failed:\n{str(e)}')
