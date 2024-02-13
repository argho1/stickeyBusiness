import os
import sys
import time
import openpyxl
import pandas as pd
from barcode import Code128
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfWriter, PdfReader


print()
eanno = "0796554198316"




# # Load data from Excel file
# workbook = openpyxl.load_workbook('modified_file.xlsx')
# sheet = workbook.active

# # Create PDF canvas
# pdf = canvas.Canvas('output.pdf', pagesize=A4)

# Constants for positioning barcodes on the page
x_start = 10 * mm  # Starting x position (20mm from the left margin)
y_start = 235 * mm  # Starting y position (280mm from the bottom margin)
x_increment = 0  # No horizontal spacing between barcodes in a row
y_increment = 18 * mm  # Vertical spacing between barcodes (adjusted to make them closer)
barcode_width = 90 * mm  # Barcode width
barcode_height = 16 * mm  # Reduced barcode height


# Load the Excel file into a pandas DataFrame
# Make sure to replace 'your_excel_file.xlsx' with the actual path to your Excel file
df = pd.read_excel('jio300_7feb24.xlsx', engine='openpyxl')  # Ensure you have 'openpyxl' installed for .xlsx files

# Assuming the column names you want to select are 'ColumnName1' and 'ColumnName2'
# Replace these with the actual column names from your Excel file
serial_number_columnName = 'SN'
wan_mac_columnName = 'WAN_MAC'

try:
    # Select the columns by name
    serial_number_list = df[serial_number_columnName].tolist()
    wan_mac_list = df[wan_mac_columnName].tolist()
except KeyError as ke:
    print("\033[31mColumns are not properly named.\033[0m")

no_of_barcode = len(serial_number_list)


exampleData="""
EXAMPLE DATA:-

+----------------+-------------------+
| SN             | WAN_MAC           |
+----------------+-------------------+
| RCRODBK01290001| 44:B5:9C:00:46:53 |
| RCRODBK01290002| 44:B5:9C:00:46:55 |
| RCRODBK01290003| 44:B5:9C:00:46:57 |
| RCRODBK01290004| 44:B5:9C:00:46:59 |
+----------------+-------------------+"""



print("Please have data in .xlsx format with comumn names as SN for Serial Number and WAN_MAC for WAN MAC like exaple below.")
print(exampleData)
print()





# Function to print the progress bar and estimate time of completion
def print_progress_bar(page, start_time, total_pages):
    # Current progress
    percentage = (page / total_pages) * 100
    bar_length = 50  # Length of the progress bar
    filled_length = int(bar_length * page // total_pages)
    bar = '█' * filled_length + '-' * (bar_length - filled_length)
    
    # Time calculations
    current_time = time.time()
    elapsed_time = current_time - start_time
    if page > 0:
        estimated_total_time = elapsed_time / page * total_pages
        remaining_time = estimated_total_time - elapsed_time
    else:
        remaining_time = 0  # Just to handle division by zero for the first page
    
    # Formatting elapsed and remaining time
    elapsed_time_formatted = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
    remaining_time_formatted = time.strftime("%H:%M:%S", time.gmtime(remaining_time))
    
    # Print the progress bar with elapsed time and estimated time of completion
    sys.stdout.write(f'\rPage {page}/{total_pages} |[{bar}| {percentage:.2f}% Completed. Elapsed: {elapsed_time_formatted}, Remaining: {remaining_time_formatted}')
    sys.stdout.flush()










#code to make stickrs pdf
def create_stickers(data1, data2, data3, barcodes):

    start_time = time.time()
    # Create a new PDF document
    output_pdf = "router_stickers.pdf"
    c = canvas.Canvas(output_pdf, pagesize=A4)

    # Set the size and position of the stickers
    sticker_width = 200
    sticker_height = 180
    margin = 50
    page_width, page_height = A4
    num_stickers_per_row = 2
    num_stickers_per_col = 3
    stickers_per_page = num_stickers_per_row * num_stickers_per_col
    num_pages = (len(barcodes) + stickers_per_page - 1) // stickers_per_page
    global i

    for page in range(num_pages):
        start_idx = page * stickers_per_page
        end_idx = min((page + 1) * stickers_per_page, len(barcodes))

        for idx, barcode_idx in enumerate(range(start_idx, end_idx)):
            
            # Calculate the row and column for the current sticker
            row = idx // num_stickers_per_row
            col = idx % num_stickers_per_row

            # Calculate the position for the current sticker
            sticker_x = margin + col * (sticker_width + margin)
            sticker_y = page_height - margin - (row + 1) * (sticker_height + margin)

            # Add the rounded rectangle to the PDF with white fill color
            c.setFillColor(colors.white)
            c.roundRect(sticker_x, sticker_y, sticker_width, sticker_height, 10, fill=1)

            # Set the font and font size for the text
            c.setFont("Helvetica-Bold", 10)

            # Calculate the position for the text based on the sticker size and the text length
            text_x = sticker_x + 10
            text_y = sticker_y + sticker_height - 15

            # Add the data as text inside the sticker
            c.setFillColor(colors.black)
            c.drawString(text_x, text_y, data1)
            c.drawString(text_x, text_y - 15, data2)
            c.drawString(text_x, text_y - 30, data3)

            c.setFont("Helvetica-Bold", 7)

            # Generate and add barcode below data3
            barcode_x = sticker_x + 50
            barcode_y = sticker_y + sticker_height - 80
            # sn = Code128(serial_number_list[barcode_idx], writer=ImageWriter())
            # sn_image = sn.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*4, "text_distance": 30, "quite_zone": 10})
            sn_image_filename = f"./bufferDEL/SN_barcode_{page}_{idx}.png"
            # sn_image.save(sn_image_filename)
            #barcode_x for - to move left barecode_y to - to move down
            c.drawImage(sn_image_filename, barcode_x-25, barcode_y-15, width=150, height=45)
            c.drawString(barcode_x-45, barcode_y+15, f'RSN:')
            c.drawString(barcode_x-44, barcode_y+8, f'{i}')
            
            # macid = Code128(wan_mac_list[barcode_idx], writer=ImageWriter())
            # macid_image = macid.render(writer_options={'module_width': 2, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
            macid_image_filename = f"./bufferDEL/WanMac_barcode_{page}_{idx}.png"
            # macid_image.save(macid_image_filename)
            #barcode_x for - to move left barecode_y to - to move down
            c.drawImage(macid_image_filename, barcode_x-25, barcode_y-55, width=150, height=45)
            c.drawString(barcode_x-45, barcode_y-25, 'MAC:')

            ean = Code128(eanno, writer=ImageWriter())
            ean_image = ean.render(writer_options={'module_width': 3, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
            ean_image_filename = f"./bufferDEL/ean_barcode.png"
            ean_image.save(ean_image_filename)
            #barcode_x for - to move left barecode_y to - to move down
            c.drawImage(ean_image_filename, barcode_x-25, barcode_y-95, width=150, height=45)
            c.drawString(barcode_x-45, barcode_y-65, 'EAN:')

            
            #removes barcode image file
            time.sleep(1)
            #os.remove(sn_image_filename)
            #os.remove(macid_image_filename)
            #os.remove(ean_image_filename)

            #ProgressBar
            print_progress_bar(i, start_time, no_of_barcode)
            i=i+1
        # Add a new page for the next set of stickers
        if page < num_pages - 1:
            c.showPage()

    #os.remove(ean_image_filename)
    # Save the PDF document
    c.save()

    return output_pdf

# count
i = 1
# Usage example
data1 = "Commodity: Outdoor Router"
data2 = "Model: Credo CR-3120-OD"
data3 = "Input: 48V PoE"
barcodes = serial_number_list



pdf_path = create_stickers(data1, data2, data3, barcodes)

print()
print(f"Sticker PDF created: {pdf_path}")






