import xlrd
import xlwt
from xlutils.copy import copy
import sys
import json
import os

def populate_excel(template_path, output_path, data):
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb = copy(rb)
    sheets = {s.name: i for i, s in enumerate(rb.sheets())}
    
    # Text Alignment and Gridline Enforcement Styles
    style_center = xlwt.easyxf(
        'font: name Arial, height 200;'
        'align: vert center, horiz center;'
        'borders: left thin, right thin, top thin, bottom thin;'
    )
    style_left = xlwt.easyxf(
        'font: name Arial, height 200;'
        'align: vert center, horiz left;'
        'borders: left thin, right thin, top thin, bottom thin;'
    )

    # 1. DASH BOARD MAPPING (Fully Dynamic)
    if "Dash Board" in sheets:
        ws = wb.get_sheet(sheets["Dash Board"])
        ws.write(6, 2, data.get("client_name", ""), style_left)     
        ws.write(16, 2, data.get("date", ""), style_left)           
        ws.write(17, 2, data.get("project_name", ""), style_left)   
        ws.write(18, 2, data.get("instrument", ""), style_left)     
        ws.write(19, 2, data.get("make", ""), style_left)           
        ws.write(20, 2, data.get("serial", ""), style_left)         

    # 2. SOFTWARE MAPPING (Fully Dynamic)
    if "Software" in sheets:
        ws = wb.get_sheet(sheets["Software"])
        ws.write(5, 1, data.get("date", ""), style_center)
        ws.write(5, 5, data.get("mode", ""), style_center)
        ws.write(6, 1, data.get("project_name", ""), style_center)
        ws.write(6, 5, data.get("capacity", ""), style_center)
        ws.write(7, 1, data.get("instrument", ""), style_center)
        ws.write(7, 5, data.get("range", ""), style_center)
        ws.write(8, 1, data.get("make", ""), style_center)
        ws.write(8, 5, data.get("increment", ""), style_center)
        ws.write(9, 1, data.get("serial", ""), style_center)
        ws.write(9, 5, data.get("resolution", ""), style_center)
        
        ws.write(13, 4, data.get("temp_before", ""), style_center)
        ws.write(13, 5, data.get("temp_after", ""), style_center)
        ws.write(14, 4, data.get("hum_before", ""), style_center)
        ws.write(14, 5, data.get("hum_after", ""), style_center)

    # 3. Data Logger Mapping (Precise Coordinates to avoid header collision)
    if "Data Logger" in sheets:
        ws = wb.get_sheet(sheets["Data Logger"])
        
        # Write metadata ONLY into the empty input cells next to the labels
        # Row 1 is Date, Row 2 is Ref, Row 3 is Item, Row 4 is Make, Row 5 is Serial
        ws.write(1, 2, data.get("date", ""), style_left)
        ws.write(2, 2, data.get("project_name", ""), style_left)
        ws.write(3, 2, data.get("instrument", ""), style_left)
        ws.write(4, 2, data.get("make", ""), style_left)
        ws.write(5, 2, data.get("serial", ""), style_left)
        
        # Mapping for the right side parameters
        ws.write(1, 10, data.get("mode", ""), style_center)
        ws.write(2, 10, data.get("capacity", ""), style_center)
        ws.write(3, 10, data.get("range", ""), style_center)
        ws.write(4, 10, data.get("increment", ""), style_center)
        ws.write(5, 10, data.get("resolution", ""), style_center)
        
        # Measurement Data Table (Starts strictly at index 17)
        points = data.get("points", [])
        start_row = 17 
        for i, pt in enumerate(points):
            row = start_row + i
            target = pt.get("target", 0)
            
            # Ensure we are writing to clean, empty cells
            ws.write(row, 1, target, style_center) 
            ws.write(row, 2, target, style_center) 
            ws.write(row, 3, pt.get("s1", 0), style_center) 
            ws.write(row, 4, target, style_center) 
            ws.write(row, 5, pt.get("s2", 0), style_center) 
            ws.write(row, 6, target, style_center) 
            ws.write(row, 7, pt.get("s3", 0), style_center)

    # 4. RECORDS SHEET MAPPING (Fully Dynamic)
    if "Records" in sheets:
        ws = wb.get_sheet(sheets["Records"])
        ws.write(11, 9, data.get("project_name", ""), style_left)
        ws.write(12, 9, data.get("instrument", ""), style_left)
        ws.write(13, 9, data.get("make", ""), style_left)
        ws.write(14, 9, data.get("client_name", ""), style_left)
        ws.write(17, 9, data.get("date", ""), style_left)

    # 5. CERTIFICATES MAPPING (Fully Dynamic)
    certificate_sheets = ["Certificate 5PTS", "Certificate 10PTS"]
    for cert in certificate_sheets:
        if cert in sheets:
            ws = wb.get_sheet(sheets[cert])
            ws.write(8, 7, data.get("date", ""), style_left)
            ws.write(9, 7, data.get("instrument", ""), style_left)
            ws.write(10, 7, data.get("capacity", ""), style_left)
            ws.write(11, 7, data.get("range", ""), style_left)
            ws.write(12, 7, data.get("increment", ""), style_left)
            ws.write(13, 7, data.get("make", ""), style_left)
            ws.write(14, 7, data.get("serial", ""), style_left)
            ws.write(16, 7, data.get("client_name", ""), style_left)

    wb.save(output_path)
    print(f"Excel report generated: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_bridge.py '<json_data>'")
        sys.exit(1)
        
    try:
        input_data = json.loads(sys.argv[1])
        template = "Testing Machine Software_revised (1).xls"
        output = f"reports/Report_{input_data.get('id', 'temp')}.xls"
        
        if not os.path.exists("reports"):
            os.makedirs("reports")
            
        populate_excel(template, output, input_data)
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)
