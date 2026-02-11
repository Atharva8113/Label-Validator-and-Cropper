import pdfplumber
import os

target_dir = r"c:/Users/Admin/Documents/NAGARKOT/Documentation/Label Copy"
# Pick one file to test
test_file = os.path.join(target_dir, "81154UN25.pdf")

if os.path.exists(test_file):
    with pdfplumber.open(test_file) as pdf:
        page = pdf.pages[0]
        print(f"Page size: {page.width}x{page.height}")
        
        # Collect all objects
        objects = []
        if hasattr(page, 'chars'): objects.extend(page.chars)
        if hasattr(page, 'lines'): objects.extend(page.lines)
        if hasattr(page, 'rects'): objects.extend(page.rects)
        if hasattr(page, 'images'): objects.extend(page.images)
        if hasattr(page, 'curves'): objects.extend(page.curves)
        
        if objects:
            x0 = min(obj['x0'] for obj in objects)
            top = min(obj['top'] for obj in objects)
            x1 = max(obj['x1'] for obj in objects)
            bottom = max(obj['bottom'] for obj in objects)
            
            print(f"Content BBox (plumber): x0={x0}, top={top}, x1={x1}, bottom={bottom}")
            print(f"Proposed Crop (plumber coords): Left={x0}, Top={top}, Right={x1}, Bottom={bottom}")
        else:
            print("No objects found.")
else:
    print(f"File {test_file} not found.")
