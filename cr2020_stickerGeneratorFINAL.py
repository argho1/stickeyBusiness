import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from barcode import Code128
from barcode.writer import ImageWriter

import datetime, os, sys


# Utility Functions
def print_banner(text):
        length = len(text) + 4
        line = '*' * length
        print(line)
        print(f'* {" " * (length - 4)} *')
        print(f'* {text} *')
        print(f'* {" " * (length - 4)} *')
        print(line)



def chooseFile(folder_path):
    def choose_file(files):
        while True:
            try:
                print("0. To go back to Main Menu.")
                choice = int(input("Enter the number corresponding to the file you want to choose: "))
                
                if choice == 0:
                    print_banner("Taking you back in time.")
                    sys.exit(1)
                if 1 <= choice <= len(files):
                    return f'{folder_path}{files[choice - 1]}'
                else:
                    print("Invalid choice. Please enter a valid number.")
            except ValueError:
                print("Invalid input. Please enter a number.")

    try:
        # Filter files to include only Excel files
        files = [file for file in os.listdir(folder_path) if file.endswith(('.xlsx', '.xls', '.xlsm'))]

        print_banner(f"Available Excel files in {folder_path}:")
        
        for i, file in enumerate(files, 1):
            file_path = os.path.join(folder_path, file)
            # Get the modification timestamp and format it as a date
            modified_date = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
            print(f"{i}. {file} - Last Modified: {modified_date})")
        
        print("\n")
        # Select file

        if not files:
            print("The folder contains no Excel files.")
            return None
        else:
            chosen_file = choose_file(files)
            print()
            print(f"\033[32mYou chose: {chosen_file}\033[0m\n")
            #chosen_file_location = f"{folder_path}{chosen_file}"
            return chosen_file#_location

    except KeyboardInterrupt:
        print("\nCTRL + C pressed, getting the duck outta here.\n")
        sys.exit(1)




def draw_text(c, x, y, text, font_size, font="Helvetica" ):
    c.setFont(font, font_size*2)
    c.drawString(x, y, text) 

def draw_barcode(c, x, y, data, angle, scale=1.6):
    # Generate barcode
    barcode = Code128(data, writer=ImageWriter())
    barcode = barcode.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 22*4, "text_distance": 34, "quite_zone": 10})
    barcode_path = f'./bufferDEL/barcode_{data}.png'
    barcode.save(barcode_path)

    # Save the current graphics state
    c.saveState()
    # Move the origin to the point where the image will be placed
    # Then rotate around this new origin
    c.translate(x, y)
    c.rotate(angle)

    # Draw the barcode image on the canvas
    # Scaling the dimensions to adjust the size of the barcode
    barcode_width = 20 * scale  # Scale width
    barcode_height = 67 * scale  # Scale height
    c.drawImage(barcode_path, -barcode_width, 0, barcode_height, barcode_width)

    # Restore the graphics state to avoid affecting other drawings
    c.restoreState()

# Function to draw a sticker
def draw_sticker(x, y, width, height, sn, imei, model_number):
    
    # STICKER FRAME
    c.line(x, y, x, y + 40 * mm + 3.5 * mm)  # Left border
    
    c.line(x, y, x + 49 * mm + 5.5 * mm, y)  # Button border
    
    c.line(x, y + 40 * mm + 3.5 * mm, x + 42 * mm + 5 * mm, y + 40 * mm + 3.5 * mm)  # Top border
    
    c.line(x + 49 * mm + 5.5 * mm, y , x + 49 * mm + 5.5 * mm, y + 33 * mm + 2.5 * mm)  # Right border
    
    c.line(x + 49 * mm + 5.6 * mm, y + 33 * mm + 2.3 * mm, x + 42 * mm + 5 * mm, y + 40 * mm + 3.5 * mm)  # Right corner tilted border

    #BARCODE
    draw_barcode(c, x + 1, y + 83, sn, -90)

    router_data = {
        "title"   : "Cellular Router",
        'model'   : "Model : CR2020",
        'power'   : 'Power : 9-36V / 1.5A',
        'version' : 'Version : V2A-S230E',
        "bands1"  : "Bands : LTE FDD(B1/B3/B5/B8)",
        "bands2"  : "             LTE TDD(B34/B38/B39/B40/B41)"
    }

    router_data = {
        "title"   :  "cWAN",
        'model'   : f"Model : {model_number}",
        'power'   : 'Power : 12V / 1A',
        "imei"    : f"IMEI : {imei}",
        }

    if len(router_data) == 4:
        font = 3
        align_x = x + 14 * mm # 14 default
        align_y = y + 26 * mm # 26 default
        space = 0
    
    elif len(router_data) == 6 :
        font = 0
        align_x = 0


    #TOP LINE
    c.line(align_x - 1.5 * mm, align_y + 5 * mm, x + width - 10, align_y + 5 * mm )  # Top border

    #HEAD
    text1="CREDO "
    text2="NETWORKS"
    head_font_size = font + 1.5
    draw_text(c, align_x, align_y + 7 * mm, text1, head_font_size, font='Helvetica-Bold')
    draw_text(c, align_x + 12 * mm, align_y + 7 * mm, text2, head_font_size)

    
    for key, value in router_data.items():
        if key == 'title':
            draw_text(c, align_x, space + align_y, f"{value}", font + 2 , font='Helvetica-Bold')
            space = space - 10
        elif key in ['bands1','bands2']:
            draw_text(c, align_x + 2, space + align_y - 1.55 * mm, f"{value}", font - 1, font='Helvetica-Bold')
            space = space - 10
        else:
            draw_text(c, align_x + 2, space + align_y - 1.55 * mm , f"{value}", font + 1)
            space = space - 10

    #BOTTOM LINE
    c.line(align_x - 1.5 * mm , align_y - 15 * mm, x + width - 10, align_y - 15 * mm)  # Top border


    text="CREDO NETWORKS"
    #TAIL
    draw_text(c,align_x + 1 * mm , align_y - 19 * mm , text, font + 1                                                                                                                                                                                                        , font='Helvetica-Bold')
    



    # Draw lines framing the sticker


# Load data from Excel
# excel_path = chooseFile("./data/")  # Adjust the path to your Excel file
excel_path = "./data/SN.xlsx"
df = pd.read_excel(excel_path)
df['IMEI'] = df['IMEI'].apply(lambda x: 'N/A' if pd.isna(x) or x == '' or x == 'nan' else int(x))

# Create a PDF for output
c = canvas.Canvas("stickers.pdf", pagesize=A4)
width, height = A4  # width and height of the page

print(df)
# Page and sticker dimensions in millimeters
page_width, page_height = A4
sticker_width, sticker_height = 49.2 * mm, 40 * mm
margin = 15 * mm  # Margin on each side
additional_space = 30 * mm  # Additional space between stickers

# Adjusting the number of columns
num_columns = 2

# Check if the total width exceeds the page width
total_width_needed = num_columns * (sticker_width + margin + additional_space) + margin
if total_width_needed > page_width:
    raise ValueError(f"Total width {total_width_needed/mm}mm exceeds page width {page_width/mm}mm. Reduce number of columns or sticker width.")

# Calculate the number of stickers per row and number of rows
stickers_per_row = num_columns  # same as the number of columns
rows_per_page = int((page_height - 2 * margin) / (sticker_height + 10 * mm))

# Define starting positions
start_x = margin + 20
start_y = page_height - margin - sticker_height

# Iterate over rows in the DataFrame
for index, row in df.iterrows():
    sn = row['SN']
    
    imei = row['IMEI']
    model_number = row['Model']
    # print(sn, " ", imei)
    row_num = (index // stickers_per_row) % rows_per_page
    column = index % stickers_per_row

    x = start_x + column * (sticker_width + margin + additional_space)
    y = start_y - row_num * (sticker_height + 15 * mm)

    draw_sticker(x , y, sticker_width, sticker_height, sn, imei, model_number)

    # Check if we need a new page
    if (index + 1) % (stickers_per_row * rows_per_page) == 0:
        c.showPage()
        start_y = page_height - margin - sticker_height  # Reset y position


# A4 dimensions in points
width, height = A4

# # Function to draw mm marks
# def draw_mm_marks():
#     # Draw horizontal marks
#     for x in range(0, int(width / mm)):
#         c.line(x * mm, 0, x * mm, 5 * mm)  # Bottom marks
#         c.line(x * mm, height, x * mm, height - 5 * mm)  # Top marks
#         if x % 10 == 0:  # Add text label every 10 mm
#             c.drawString(x * mm + 2, 8, str(x))
#             c.drawString(x * mm + 2, height - 12, str(x))
    
#     # Draw vertical marks
#     for y in range(0, int(height / mm)):
#         c.line(0, y * mm, 5 * mm, y * mm)  # Left marks
#         c.line(width, y * mm, width - 5 * mm, y * mm)  # Right marks
#         if y % 10 == 0:  # Add text label every 10 mm
#             c.drawString(8, y * mm + 2, str(y))
#             c.drawString(width - 20, y * mm + 2, str(y))

# # Draw mm marks on the canvas
# draw_mm_marks()

# Save the PDF
c.save()


