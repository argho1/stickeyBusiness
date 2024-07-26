import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from barcode import Code128
from barcode.writer import ImageWriter


# Load data from Excel
excel_path = 'SN.xlsx'  # Adjust the path to your Excel file
df = pd.read_excel(excel_path)

# Create a PDF for output
c = canvas.Canvas("stickers.pdf", pagesize=A4)
width, height = A4  # width and height of the page

def draw_text(c, x, y, text, font_size, font="Helvetica" ):
    c.setFont(font, font_size*2)
    c.drawString(x, y, text) 

def draw_barcode(c, x, y, data, angle, scale=1.3):
    # Generate barcode
    barcode = Code128(data, writer=ImageWriter())
    barcode = barcode.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 22*4, "text_distance": 34, "quite_zone": 10})
    barcode_path = f'barcode_{data}.png'
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

    scale_x = -15
    scale_y = -7
    # router_data = {
    #     "title"   :  "Cellular Router",
    #     'model'   : "Model : CR2020",
    #     'power'   : 'Power : 9-36V / 1.5A',
    #     'version' : 'Version : V2A-S230E',
    #     "bands1"  : "Bands : LTE FDD(B1/B3/B5/B8)",
    #     "bands2"  : "             LTE TDD(B34/B38/B39/B40/B41)"
    # }

    if imei == 0:
        imei = "N/A"

    router_data = {
        "title"   :  "cWAN",
        'model'   : f"Model : {model_number}",
        'power'   : 'Power : 12V / 1A',
        "imei"    : f"IMEI : {imei}",
        }

    #BARCODE
    draw_barcode(c, x- 12, y+56, sn, -90)

    text1="CREDO "
    text2="NETWORKS"
    #HEAD
    draw_text(c, x+35 + scale_x, y+73 + scale_y, text1, 8, font='Helvetica-Bold')
    draw_text(c, x+35+60 + scale_x, y+73 + scale_y, text2, 8)

    space = y-17
    for key, value in router_data.items():
        if key == 'title':
            draw_text(c, x+35 + scale_x, space+73 + scale_y, f"{value}", 5, font='Helvetica-Bold')
            space = space - 10
        else:
            draw_text(c, x+37 + scale_x , space+73 + scale_y, f"{value}", 4)
            space = space - 11

    text="CREDO NETWORKS"
    #TAIL
    draw_text(c, x+58 + scale_x, y+5 + scale_y, text, 5, font='Helvetica-Bold')
    



    # Draw lines framing the sticker
    #HEAD LINE
    c.line(x+30 + scale_x, y+15 + scale_y, x + width, y+15 + scale_y)  # Top border
    #TAIL LINE
    c.line(x+30 + scale_x, y+70 + scale_y, x + width, y+70 + scale_y)  # Top border


# Page and sticker dimensions in millimeters
page_width, page_height = A4
sticker_width, sticker_height = 63.2 * mm, 30 * mm
margin = 10 * mm  # Margin on each side

# Adjusting the number of columns
num_columns = 2

# Check if the total width exceeds the page width
total_width_needed = num_columns * (sticker_width + margin) + margin
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
    imei = int(row['IMEI'])
    model_number = row['Model']
    print(sn, " ", imei)
    row_num = (index // stickers_per_row) % rows_per_page
    column = index % stickers_per_row

    x = start_x + column * (sticker_width + margin)
    y = start_y - row_num * (sticker_height + 10 * mm)

    draw_sticker(x , y, sticker_width, sticker_height, sn, imei, model_number)

    # Check if we need a new page
    if (index + 1) % (stickers_per_row * rows_per_page) == 0:
        c.showPage()
        start_y = page_height - margin - sticker_height  # Reset y position

# Save the PDF
c.save()
