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
from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter, A4, landscape













def router_box_stickers():



    #code to make stickrs pdf
    def create_stickers(data1, data2, data3, barcodes):

        start_time = time.time()

        # Create a new PDF document
        output_pdf = "router_box_stickers.pdf"
        c = canvas.Canvas(output_pdf, pagesize=landscape(A4))


        # Set the size and position of the stickers
        sticker_width = 264.922 * mm
        sticker_height = 82.63 * mm
        margin = 50
        page_width, page_height = landscape(A4)
        num_stickers_per_row = 1
        num_stickers_per_col = 2
        stickers_per_page = num_stickers_per_row * num_stickers_per_col
        num_pages = (len(barcodes) + stickers_per_page - 1) // stickers_per_page
        i = 1

        for page in range(num_pages):
            start_idx = page * stickers_per_page
            end_idx = min((page + 1) * stickers_per_page, len(barcodes))

            for idx, barcode_idx in enumerate(range(start_idx, end_idx)):
                
                # Calculate the row and column for the current sticker
                row = idx // num_stickers_per_row
                col = idx % num_stickers_per_row

                # Calculate the position for the current sticker
                sticker_x = margin + col * (sticker_width + margin)
                sticker_y = page_height - margin - (row + 1) * (sticker_height + margin) + 80

                # Add the rounded rectangle to the PDF with white fill color
                c.setFillColor(colors.white)
                c.roundRect(sticker_x, sticker_y, sticker_width, sticker_height, 10, fill=1)

                # Set the font and font size for the text
                c.setFont("Helvetica-Bold", 15)

                # Calculate the position for the text based on the sticker size and the text length
                text_x = (sticker_x + 10 ) 
                text_y = (sticker_y - 20 + sticker_height - 15) 

                # Add the data as text inside the sticker
                c.setFillColor(colors.black)
                c.drawString(text_x, text_y, data1)
                c.drawString(text_x+150, text_y, data1a)
                c.drawString(text_x, text_y - 20, data2)
                c.drawString(text_x+150, text_y - 20, data2a)
                c.drawString(text_x, text_y - 40, data3)
                c.drawString(text_x+150, text_y - 40, data3a)
                c.drawString(text_x+160, text_y - 60, data31)
                c.drawString(text_x, text_y - 85, data4)
                c.drawString(text_x, text_y - 120, data5)
                c.drawString(text_x+150, text_y - 120, data5a)
                c.drawString(text_x+160, text_y - 140, data51)
                c.drawString(text_x, text_y - 160, data6)
                c.drawString(text_x+150, text_y - 160, data6a)
                c.drawString(text_x, text_y - 180, data7)
                c.drawString(text_x+150, text_y - 180, data7a)

                # Generate and add barcode below data3
                barcode_x = sticker_x * 10.8 + 50
                barcode_y = sticker_y + sticker_height - 80

                
                rcno = Code128(sn_list[barcode_idx], writer=ImageWriter())
                rcno_image = rcno.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*4, "text_distance": 30, "quite_zone": 10})
                rcno_image_filename = f"barcode_{page}_{idx}.png"
                rcno_image.save(rcno_image_filename)
                #barcode_x for - to move left barecode_y to - to move down
                c.drawImage(rcno_image_filename, barcode_x-25, barcode_y-10, width=150+70, height=55)
                c.drawString(barcode_x-70, barcode_y+18, f"RSN :")
                c.drawString(barcode_x-65, barcode_y+5, f"{i}")


                ean = Code128(eanno, writer=ImageWriter())
                ean_image = ean.render(writer_options={'module_width': 4, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
                ean_image_filename = f"ean_barcode.png"
                ean_image.save(ean_image_filename)
                #barcode_x for - to move left barecode_y to - to move down
                c.drawImage(ean_image_filename, barcode_x-25, barcode_y-70, width=150+70, height=55)
                c.drawString(barcode_x-70, barcode_y-40, "EAN :")


                
                macid = Code128(wan_mac_list[barcode_idx], writer=ImageWriter())
                macid_image = macid.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
                macid_image_filename = f"barcode1_{page}_{idx}.png"
                macid_image.save(macid_image_filename)
                #barcode_x for - to move left barecode_y to - to move down
                c.drawImage(macid_image_filename, barcode_x-25, barcode_y-130, width=150+70, height=55)
                c.drawString(barcode_x-70, barcode_y-100, "MAC :")
                



                #removes barcode image file
                os.remove(rcno_image_filename)
                os.remove(macid_image_filename)
                os.remove(ean_image_filename)

                #ProgressBar
                print_progress_bar(i, start_time, no_of_barcode)
                i=i+1
            # Add a new page for the next set of stickers
            if page < num_pages - 1:
                c.showPage()

        # Save the PDF document
        c.save()

        return output_pdf




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





    # # Load data from Excel file
    # workbook = openpyxl.load_workbook('input.xlsx')
    # sheet = workbook.active

    # Constants for positioning barcodes on the page
    x_start = 10 * mm  # Starting x position (20mm from the left margin)
    y_start = 235 * mm  # Starting y position (280mm from the bottom margin)
    x_increment = 0  # No horizontal spacing between barcodes in a row
    y_increment = 18 * mm  # Vertical spacing between barcodes (adjusted to make them closer)
    barcode_width = 90 * mm  # Barcode width
    barcode_height = 16 * mm  # Reduced barcode height


    # Load the Excel file into a pandas DataFrame
    # Make sure to replace 'your_excel_file.xlsx' with the actual path to your Excel file
    df = pd.read_excel('modified_file.xlsx', engine='openpyxl')  # Ensure you have 'openpyxl' installed for .xlsx files

    # Assuming the column names you want to select are 'ColumnName1' and 'ColumnName2'
    # Replace these with the actual column names from your Excel file
    column_name_1 = 'SN'
    column_name_2 = 'WAN_MAC'

    try:
        # Select the columns by name
        sn_list = df[column_name_1].tolist()
        wan_mac_list = df[column_name_2].tolist()
    except KeyError as ke:
        print("\033[31mColumns are not properly named.\033[0m")

    no_of_barcode = len(sn_list)


    # Static Data
    data1 = 'Commodity'
    data1a =': Credo CR-3120-OD Router'
    data2 = 'Manufactured By'
    data2a =': Tenet Networks Private Limited'
    data3 = 'Net Quantity'
    data3a=': 1 Outdoor Router + 1 Patch Cord'
    data31= ' + 1 POE Adapter + 1 clamp'
    data4 = 'Month & Year of Manufacture: 02/2024'
    data5 = 'Office Address'
    data5a=': A-541, Logix Technova Sector-132'
    data51= 'Noida-201305 U.P. India'
    data6 = 'Customer Care No.'
    data6a=': +91 120-4165905'
    data7 = 'Email ID'
    data7a=': info@tenetnetworks.com'
    barcodes = sn_list

    #Iteration Count
    #i = 1


    #ESN no.
    eanno =("0796554198316")
    print()


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
    pdf_path = create_stickers(data1, data2, data3, barcodes)

    print()
    print(f"Sticker PDF created: {pdf_path}")




    #'module_width': 10, 'module_height': 80, "font_size": 20*5, "text_distance": 28


router_box_stickers()