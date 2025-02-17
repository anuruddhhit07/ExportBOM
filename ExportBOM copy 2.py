import adsk.core, adsk.fusion, traceback, time, csv, os, base64,shutil

# Global variable to control whether volume is included
includeVolume = False  # Set this to False if you do not want to include volume in the BOM

def spacePadRight(value, length):
    pad = ''
    if isinstance(value, str):
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
        mStr += spacePadRight(item['name'], 25) + str(spacePadRight(item['instances'], 15)) + str(spacePadRight(item['mat'], 15))
        # Include volume only if it's part of the BOM
        if includeVolume:
            mStr += str(item['volume'])
        mStr += '\n'
    return mStr
def delete_related_files(filename):
    """
    Deletes all files and folders related to the given filename.

    Parameters:
        filename (str): The full path to the base file.
    """
    # Get the base directory and name without extension
    base_dir = os.path.dirname(filename)
    base_name = os.path.splitext(filename)[0]
    
    # Define the directory that needs to be deleted (e.g., '_files' folder)
    related_dir = f"{base_name}_files"
    
    # Delete the related directory if it exists
    if os.path.exists(related_dir) and os.path.isdir(related_dir):
        shutil.rmtree(related_dir)
        print(f"Deleted directory: {related_dir}")
    
    # Optionally delete any files with related names or extensions
    for file in os.listdir(base_dir):
        if file.startswith(base_name):
            file_path = os.path.join(base_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
def buildCSV0(bom, imageDirectory, fileName):
    with open(fileName + '.csv', 'w', encoding='utf-8', newline='') as csvfile:
        writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, delimiter=',')
        writer.writerow(['name', 'instances', 'mat', 'volume' if includeVolume else '', 'image_base64'])
        for item in bom:
            name = item['name']
            # Encode the image as base64
            image_path = os.path.join(imageDirectory, f"{item['component'].id}.png")
            image_base64 = encode_image_to_base64(image_path) if os.path.exists(image_path) else ''
            # Write volume only if included
            if includeVolume:
                writer.writerow([name, item['instances'], item['mat'], item['volume'], image_base64])
            else:
                writer.writerow([name, item['instances'], item['mat'], image_base64])

def buildCSV(bom, imageDirectory, fileName):
    with open(fileName + '.csv', 'w', encoding='utf-8', newline='') as csvfile:
        writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, delimiter=',')
        # Write header row, including volume if necessary
        if includeVolume:
            writer.writerow(['name', 'instances', 'mat', 'volume', 'image_path'])
        else:
            writer.writerow(['name', 'instances', 'mat', 'image_path'])
        
        for item in bom:
            name = item['name']
            # Construct the image path
            image_path = os.path.join(imageDirectory, f"{item['component'].id}.png")
            # Write volume only if included
            if includeVolume:
                writer.writerow([name, item['instances'], item['mat'], item['volume'], image_path])
            else:
                writer.writerow([name, item['instances'], item['mat'], image_path])
            
def encode_image_to_base64(image_path):
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Error encoding image {image_path}: {str(e)}")
        return ''

def buildHTMLWithImages(app, bom, imageDirectory, fileName):
    with open(fileName + '.html', 'w', encoding='utf-8') as html_file:
        html_file.write('<html><body>\n')
        html_file.write('<table border="1">\n')
        
        # Conditionally write the header with or without the volume column
        html_file.write('<tr><th>Name</th><th>Instances</th><th>Material</th>')
        if includeVolume:
            html_file.write('<th>Volume</th>')  # Only include this if includeVolume is True
        html_file.write('<th>Image</th></tr>\n')
        
        # Write the rows
        for item in bom:
            image_path = f"{imageDirectory}/{item['component'].id}.png"
            image_base64 = encode_image_to_base64(image_path) if os.path.exists(image_path) else ''
            html_file.write('<tr>')
            html_file.write(f'<td>{item["name"]}</td>')
            html_file.write(f'<td>{item["instances"]}</td>')
            html_file.write(f'<td>{item["mat"]}</td>')
            
            # Include volume data only if includeVolume is True
            if includeVolume:
                html_file.write(f'<td>{item["volume"]}</td>')
            
            if image_base64:
                html_file.write(f'<td><img src="data:image/png;base64,{image_base64}" alt="Image" width="50" height="50"></td>')
            else:
                html_file.write('<td>Image not found</td>')
            
            html_file.write('</tr>\n')
        
        html_file.write('</table>\n')
        html_file.write('</body></html>\n')

def buildHTMLWithImagesEditable0(app, bom, imageDirectory, fileName):
    with open(fileName+'_editable' + '.html', 'w', encoding='utf-8') as html_file:
        html_file.write('<html><body>\n')
        html_file.write('<table border="1">\n')
        
        # Conditionally write the header with or without the volume column
        html_file.write('<tr><th>Name</th><th>Instances</th><th>Material</th>')
        if includeVolume:
            html_file.write('<th>Volume</th>')  # Only include this if includeVolume is True
        html_file.write('<th>Image</th></tr>\n')
        
        # Write the rows with editable cells
        for item in bom:
            image_path = f"{imageDirectory}/{item['component'].id}.png"
            image_base64 = encode_image_to_base64(image_path) if os.path.exists(image_path) else ''
            html_file.write('<tr>')
            
            # Editable fields (contenteditable="true")
            html_file.write(f'<td contenteditable="true">{item["name"]}</td>')
            html_file.write(f'<td contenteditable="true">{item["instances"]}</td>')
            html_file.write(f'<td contenteditable="true">{item["mat"]}</td>')
            
            # Include volume data only if includeVolume is True
            if includeVolume:
                html_file.write(f'<td contenteditable="true">{item["volume"]}</td>')
            
            if image_base64:
                html_file.write(f'<td><img src="data:image/png;base64,{image_base64}" alt="Image" width="50" height="50"></td>')
            else:
                html_file.write('<td>Image not found</td>')
            
            html_file.write('</tr>\n')
        
        html_file.write('</table>\n')
        html_file.write('</body></html>\n')

def buildHTMLWithImagesEditableCSV(app, bom, imageDirectory, fileName):
    with open(fileName + '_editable' + '.html', 'w', encoding='utf-8') as html_file:
        html_file.write('<html><body>\n')
        
        # Include a script for handling selection and exporting
        html_file.write('''
        <script>
        function exportSelected() {
            var table = document.getElementById("bomTable");
            var selectedRows = [];
            var checkboxes = table.getElementsByClassName("rowCheckbox");
            for (var i = 0; i < checkboxes.length; i++) {
                if (checkboxes[i].checked) {
                    var row = checkboxes[i].closest("tr");
                    var rowData = [];
                    var cells = row.getElementsByTagName("td");
                    for (var j = 1; j < cells.length - 1; j++) { // Skip the image cell
                        rowData.push(cells[j].innerText);
                    }
                    selectedRows.push(rowData);
                }
            }

            if (selectedRows.length > 0) {
                var csvContent = "Name,Instances,Material";
                if (''' + ('true' if includeVolume else 'false') + ''') {
                    csvContent += ",Volume";  // Include Volume if enabled
                }
                csvContent += ",Image\\n";  // Include Image column for the export
                
                selectedRows.forEach(function(row) {
                    csvContent += row.join(",") + "\\n";  // Combine the row data
                });
                
                var blob = new Blob([csvContent], { type: "text/csv" });
                var link = document.createElement("a");
                link.href = URL.createObjectURL(blob);
                link.download = "selected_bom.csv";  // Default export file name
                link.click();
            } else {
                alert("No rows selected for export.");
            }
        }
        </script>
        ''')

        html_file.write('<table id="bomTable" border="1">\n')
        
        # Conditionally write the header with or without the volume column
        html_file.write('<tr><th>Select</th><th>Name</th><th>Instances</th><th>Material</th>')
        if includeVolume:
            html_file.write('<th>Volume</th>')  # Only include this if includeVolume is True
        html_file.write('<th>Image</th></tr>\n')
        
        # Write the rows with editable cells and selection checkboxes
        for item in bom:
            image_path = f"{imageDirectory}/{item['component'].id}.png"
            image_base64 = encode_image_to_base64(image_path) if os.path.exists(image_path) else ''

            html_file.write('<tr>')
            
            # Selection checkbox
            html_file.write(f'<td><input type="checkbox" class="rowCheckbox" checked="checked"></td>')
            
            # Editable fields (contenteditable="true")
            html_file.write(f'<td contenteditable="true">{item["name"]}</td>')
            html_file.write(f'<td contenteditable="true">{item["instances"]}</td>')
            html_file.write(f'<td contenteditable="true">{item["mat"]}</td>')
            
            # Include volume data only if includeVolume is True
            if includeVolume:
                html_file.write(f'<td contenteditable="true">{item["volume"]}</td>')
            
            # Handling the image column - export either base64 or image URL
            if image_base64:
                html_file.write(f'<td><img src="data:image/png;base64,{image_base64}" alt="Image" width="50" height="50"></td>')
            else:
                html_file.write('<td>Image not found</td>')
            
            html_file.write('</tr>\n')
        
        html_file.write('</table>\n')
        
        # Export button to trigger CSV export
        html_file.write('<button onclick="exportSelected()">Export Selected Rows</button>\n')
        
        html_file.write('</body></html>\n')


def buildHTMLWithImagesEditableXlsx(app, bom, imageDirectory, fileName):
    with open(fileName+'_editable' + '.html', 'w', encoding='utf-8') as html_file:
        html_file.write('<html><body>\n')
        
        # Include a script for handling selection and exporting to XLSX
        html_file.write('''
        <script>
        function exportSelected() {
            var table = document.getElementById("bomTable");
            var selectedRows = [];
            var checkboxes = table.getElementsByClassName("rowCheckbox");
            for (var i = 0; i < checkboxes.length; i++) {
                if (checkboxes[i].checked) {
                    var row = checkboxes[i].closest("tr");
                    var rowData = [];
                    var cells = row.getElementsByTagName("td");
                    for (var j = 0; j < cells.length - 1; j++) { // Skip the image cell
                        rowData.push(cells[j].innerText);
                    }
                    selectedRows.push(rowData);
                }
            }

            if (selectedRows.length > 0) {
                var data = {
                    rows: selectedRows,
                    includeVolume: ''' + ('true' if includeVolume else 'false') + '''
                };

                fetch('/export_xlsx', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(data)
                })
                .then(response => response.blob())
                .then(blob => {
                    var link = document.createElement('a');
                    link.href = URL.createObjectURL(blob);
                    link.download = 'selected_bom.xlsx';
                    link.click();
                })
                .catch(error => alert('Error exporting data: ' + error));
            } else {
                alert("No rows selected for export.");
            }
        }
        </script>
        ''')

        html_file.write('<table id="bomTable" border="1">\n')
        
        # Conditionally write the header with or without the volume column
        html_file.write('<tr><th>Select</th><th>Name</th><th>Instances</th><th>Material</th>')
        if includeVolume:
            html_file.write('<th>Volume</th>')  # Only include this if includeVolume is True
        html_file.write('<th>Image</th></tr>\n')
        
        # Write the rows with editable cells and selection checkboxes
        for item in bom:
            image_path = f"{imageDirectory}/{item['component'].id}.png"
            image_base64 = encode_image_to_base64(image_path) if os.path.exists(image_path) else ''
            html_file.write('<tr>')
            
            # Selection checkbox
            html_file.write(f'<td><input type="checkbox" class="rowCheckbox"></td>')
            
            # Editable fields (contenteditable="true")
            html_file.write(f'<td contenteditable="true">{item["name"]}</td>')
            html_file.write(f'<td contenteditable="true">{item["instances"]}</td>')
            html_file.write(f'<td contenteditable="true">{item["mat"]}</td>')
            
            # Include volume data only if includeVolume is True
            if includeVolume:
                html_file.write(f'<td contenteditable="true">{item["volume"]}</td>')
            
            if image_base64:
                html_file.write(f'<td><img src="data:image/png;base64,{image_base64}" alt="Image" width="50" height="50"></td>')
            else:
                html_file.write('<td>Image not found</td>')
            
            html_file.write('</tr>\n')
        
        html_file.write('</table>\n')
        
        # Export button
        html_file.write('<button onclick="exportSelected()">Export Selected Rows</button>\n')
        
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

def takeImage(app, ui, component, occs, path):
    cameraTarget = False
    occurrence = False

    for occ in occs:
        comp = occ.component

        if comp == component and not cameraTarget:
            cameraTarget = adsk.core.Point3D.create(occ.transform.translation.x, occ.transform.translation.y, occ.transform.translation.z)
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
        camera.eye = adsk.core.Point3D.create(100 + cameraTarget.x, -100 + cameraTarget.y, 100 + cameraTarget.z)
        app.activeViewport.camera = camera
        app.activeViewport.refresh()
        adsk.doEvents()

        # Save the image
        image_path = os.path.join(path, f"{component.id}.png")
        success = app.activeViewport.saveAsImageFile(image_path, 128, 128)
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
        fileDialog.title = "Save BOM As"
        fileDialog.filter = 'HTML (*.html)'
        fileDialog.initialFilename = product.rootComponent.name
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showSave()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
            delete_related_files(filename)
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
            takeImage(app, ui, bomItem['component'], occs, dst_directory)
            Unisolate(visibleTopLevelComp)

        # Display BOM data and save
        msg = spacePadRight('Name', 25) + spacePadRight('Instances', 15) + spacePadRight('Material', 15) + 'Volume\n' + walkThrough(bom)
        buildCSV(bom, dst_directory, os.path.splitext(filename)[0] + '_bom')
        buildHTMLWithImages(app, bom, dst_directory, os.path.splitext(filename)[0] + '_html')
        buildHTMLWithImagesEditableCSV(app, bom, dst_directory, os.path.splitext(filename)[0] + '_html')
        ui.messageBox(msg, 'Bill Of Materials')

    except Exception as e:
        if ui:
            ui.messageBox(f'Failed:\n{str(e)}')
