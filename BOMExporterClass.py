import adsk.core, adsk.fusion, traceback, time, csv, os, base64, shutil



class BOMExporter:
    def __init__(self, include_volume=False):
        """
        Initialize the BOMExporter.

        :param include_volume: Whether to include the volume in the BOM export.
        """
        self.include_volume = include_volume

    @staticmethod
    def space_pad_right(value, length):
        pad = ''
        if isinstance(value, str):
            padding_length = length - len(value) + 1
        else:
            padding_length = length - value + 1
        while padding_length > 0:
            pad += ' '
            padding_length -= 1
        return str(value) + pad

    def walk_through(self, bom):
        result = ''
        for item in bom:
            result += self.space_pad_right(item['name'], 25) + str(self.space_pad_right(item['instances'], 15)) + str(self.space_pad_right(item['mat'], 15))
            if self.include_volume:
                result += str(item['volume'])
            result += '\n'
        return result

    @staticmethod
    def delete_related_files(filename):
        base_dir = os.path.dirname(filename)
        base_name = os.path.splitext(filename)[0]
        related_dir = f"{base_name}_files"
        
        if os.path.exists(related_dir) and os.path.isdir(related_dir):
            shutil.rmtree(related_dir)
            print(f"Deleted directory: {related_dir}")
        
        for file in os.listdir(base_dir):
            if file.startswith(base_name):
                file_path = os.path.join(base_dir, file)
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")

    def build_csv(self, bom, image_directory, file_name):
        with open(file_name + '.csv', 'w', encoding='utf-8', newline='') as csvfile:
            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC, delimiter=',')
            base_url = 'www.example.com/images/'
            if self.include_volume:
                writer.writerow(['name', 'instances', 'mat', 'volume', 'image_path'])
            else:
                writer.writerow(['name', 'instances', 'mat', 'image_path'])

            for item in bom:
                name = item['name']
                image_path = os.path.join(image_directory, f"{item['component'].id}.png")
                image_url = f"{image_directory}/{item['component'].id}.png"
                image_hyperlink = f'=HYPERLINK("{image_path}", "Open Image")'
                if self.include_volume:
                    writer.writerow([name, item['instances'], item['mat'], item['volume'], image_hyperlink])
                else:
                    writer.writerow([name, item['instances'], item['mat'], image_hyperlink])

    @staticmethod
    def encode_image_to_base64(image_path):
        try:
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode('utf-8')
        except Exception as e:
            print(f"Error encoding image {image_path}: {str(e)}")
            return ''

    def build_html_with_images(self, app, bom, image_directory, file_name, editable=False):
        suffix = '_editable' if editable else ''
        with open(file_name + suffix + '.html', 'w', encoding='utf-8') as html_file:
            html_file.write('<html><body>\n')
            html_file.write('<table border="1">\n')
            
            html_file.write('<tr><th>Name</th><th>Instances</th><th>Material</th>')
            if self.include_volume:
                html_file.write('<th>Volume</th>')
            html_file.write('<th>Image</th></tr>\n')

            for item in bom:
                image_path = f"{image_directory}/{item['component'].id}.png"
                image_base64 = self.encode_image_to_base64(image_path) if os.path.exists(image_path) else ''
                html_file.write('<tr>')

                if editable:
                    html_file.write(f'<td contenteditable="true">{item["name"]}</td>')
                    html_file.write(f'<td contenteditable="true">{item["instances"]}</td>')
                    html_file.write(f'<td contenteditable="true">{item["mat"]}</td>')
                else:
                    html_file.write(f'<td>{item["name"]}</td>')
                    html_file.write(f'<td>{item["instances"]}</td>')
                    html_file.write(f'<td>{item["mat"]}</td>')
                
                if self.include_volume:
                    html_file.write(f'<td>{item["volume"]}</td>')

                if image_base64:
                    html_file.write(f'<td><img src="data:image/png;base64,{image_base64}" alt="Image" width="50" height="50"></td>')
                else:
                    html_file.write('<td>Image not found</td>')

                html_file.write('</tr>\n')

            html_file.write('</table>\n')
            html_file.write('</body></html>\n')


    def buildHTMLWithImagesEditableCSV(self, app, bom, image_directory, file_name, editable=True):
        with open(file_name + '_editable' + '.html', 'w', encoding='utf-8') as html_file:
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
                    if (''' + ('true' if self.include_volume else 'false') + ''') {
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
            if self.include_volume:
                html_file.write('<th>Volume</th>')  # Only include this if includeVolume is True
            html_file.write('<th>Image</th></tr>\n')
            
            # Write the rows with editable cells and selection checkboxes
            for item in bom:
                image_path = f"{image_directory}/{item['component'].id}.png"
                image_base64 = self.encode_image_to_base64(image_path) if os.path.exists(image_path) else ''

                html_file.write('<tr>')
                
                # Selection checkbox
                html_file.write(f'<td><input type="checkbox" class="rowCheckbox" checked="checked"></td>')
                
                # Editable fields (contenteditable="true")
                html_file.write(f'<td contenteditable="true">{item["name"]}</td>')
                html_file.write(f'<td contenteditable="true">{item["instances"]}</td>')
                html_file.write(f'<td contenteditable="true">{item["mat"]}</td>')
                
                # Include volume data only if includeVolume is True
                if self.include_volume:
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


    def take_image(self, app, ui, component, occs, path):
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

            self.setGridDisplay(False)

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

    @staticmethod
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
    
    @staticmethod
    def Unisolate(occs):
        for occ in occs:
            occ.isLightBulbOn = True
    