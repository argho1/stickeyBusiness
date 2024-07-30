import os
import sys
import time
import shutil
import openpyxl
import datetime
import pandas as pd
from barcode import Code128
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter, A4, landscape


##Defining colors for output.##
# Normal colors
RED='\033[0;31m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[0;33m'

# Brighter shades
BRIGHT_RED='\033[1;31m'
BRIGHT_GREEN='\033[1;32m'
BRIGHT_BLUE='\033[1;34m'
BRIGHT_YELLOW='\033[1;33m'

# No Color
CEND='\033[0m'  



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







def router_body_stickers():
    print()

        #code to make stickrs pdf
    def create_stickers(selected_template, barcodes):

        commodity_text_print = selected_template['commodity_text']
        model_text_print = selected_template['model_text']
        input_text_print = selected_template['input_text']

        try:
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
                    c.drawString(text_x, text_y, f"Commodity : {commodity_text_print}")
                    c.drawString(text_x, text_y - 15, f"Model : {model_text_print}")
                    c.drawString(text_x, text_y - 30, f"Input : {input_text_print}")

                    c.setFont("Helvetica-Bold", 7)

                    # Generate and add barcode below input_text_print
                    barcode_x = sticker_x + 50
                    barcode_y = sticker_y + sticker_height - 80
                    sn = Code128(str(serial_number_list[barcode_idx]), writer=ImageWriter())
                    sn_image = sn.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*4, "text_distance": 30, "quite_zone": 10})
                    sn_image_filename = f"./bufferDEL/SN_barcode_{serial_number_list[barcode_idx]}.png"
                    sn_image.save(sn_image_filename)
                    #barcode_x for - to move left barecode_y to - to move down
                    c.drawImage(sn_image_filename, barcode_x-25, barcode_y-35, width=150, height=45)
                    c.drawString(barcode_x-45, barcode_y-5, f'{EXCEL_COLUMN_1_NAME}:')
                    # c.drawString(barcode_x-44, barcode_y-10, f'{i}')
                    
                    macid = Code128(str(wan_mac_list[barcode_idx]), writer=ImageWriter())
                    macid_image = macid.render(writer_options={'module_width': 2, 'module_height': 80, "font_size": 30*2, "text_distance": 25, "quite_zone": 10})
                    macid_image_filename = f"./bufferDEL/WanMac_barcode_{wan_mac_list[barcode_idx]}.png"
                    macid_image.save(macid_image_filename)
                    # barcode_x for - to move left barecode_y to - to move down
                    c.drawImage(macid_image_filename, barcode_x-25, barcode_y-85, width=150, height=45)
                    c.drawString(barcode_x-45, barcode_y-57, f'{EXCEL_COLUMN_2_NAME}:')

                    # ean = Code128(eanno, writer=ImageWriter())
                    # ean_image = ean.render(writer_options={'module_width': 3, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
                    # ean_image_filename = f"./bufferDEL/ean_barcode.png"
                    # ean_image.save(ean_image_filename)
                    # #barcode_x for - to move left barecode_y to - to move down
                    # c.drawImage(ean_image_filename, barcode_x-25, barcode_y-95, width=150, height=45)
                    # c.drawString(barcode_x-45, barcode_y-65, 'EAN:')

                    
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
        
        except FileNotFoundError:
            
            os.makedirs(directory)
            create_stickers(select_template, serial_number_list)
            shutil.rmtree(directory)
        
        except KeyboardInterrupt:
            print()
            print_banner("Rage Quit Inititated!! Deleteing ./bufferDEL folder.")
            shutil.rmtree(directory)
            sys.exit(1)





    def validateExcel(chosen_excel_file, EXCEL_COLUMN_1_NAME, EXCEL_COLUMN_2_NAME):
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(chosen_excel_file, engine='openpyxl')


        exampleData="""\033[33m
            EXAMPLE DATA:-

            +----------------+-------------------+
            | SN             | WAN_MAC           |
            +----------------+-------------------+
            | RCRODBK01290001| 44:B5:9C:00:46:53 |
            | RCRODBK01290002| 44:B5:9C:00:46:55 |
            | RCRODBK01290003| 44:B5:9C:00:46:57 |
            | RCRODBK01290004| 44:B5:9C:00:46:59 |
            +----------------+-------------------+\033[0m"""

        try:
            # Select the columns by name
            serial_number_list = df[EXCEL_COLUMN_1_NAME].tolist()
            wan_mac_list = df[EXCEL_COLUMN_2_NAME].tolist()
            
        except KeyError as ke:
            print("\033[31mColumns are not properly named.\033[0m")
            print()
            print("\033[33mPlease have data in .xlsx format with comumn names as SN for Serial Number and WAN_MAC for WAN MAC like example below.\033[0m")
            print(exampleData)
            return




        if len(serial_number_list) == len(wan_mac_list) and df[EXCEL_COLUMN_1_NAME].isnull().sum() == 0 and df[EXCEL_COLUMN_2_NAME].isnull().sum() == 0:
            print()
            return serial_number_list, wan_mac_list
        else:
            print()
            print("\x1b[31mValue missing or MAC vs Serial Number count mismatch!!\x1b[0m")
            print("\x1b[31mPlease Check Excel Sheet and Try Again!!\x1b[0m")
            print("\033[31mColumns are not properly named.\033[0m")
            print()
            print("\033[33mPlease have data in .xlsx format with comumn names as SN for Serial Number and WAN_MAC for WAN MAC like example below.\033[0m")
            print(exampleData)
            return
        


    def get_custom_input(selected_template):

        # print_banner("here")
        # print(selected_template)
        for key, value in selected_template.items():
            print(key, ":", value)





        return 
    
    def select_template(template_choice):
        templates = {
            # ODCP
            '1': {
                "commodity_text": "Outdoor Router",
                "model_text": "Credo CR-3120-OD",
                "input_text": "48V PoE",
                "eanno": "0796554198316"
                },
            # cWAN
            '2': {
                "commodity_text": "CWAN",
                "model_text": "CR3181-X",
                "input_text": "240V AC",
                },
            # CR2020
            '3': {
                "title"   : "Cellular Router",
                'model'   : "Model : CR2020",
                'power'   : 'Power : 9-36V / 1.5A',
                'version' : 'Version : V2A-S230E',
                "bands1"  : "Bands : LTE FDD(B1/B3/B5/B8)",
                "bands2"  : "       LTE TDD(B34/B38/B39/B40/B41)"
            },
            # cWAN Black Box
            '4': {
                "title"   :  "cWAN",
                'model'   : "Model : CR1112-A",
                'power'   : 'Power : 12V / 1A',
                "imei"    : "IMEI : # FROM EXCEL #",
            },
        }

        if template_choice in templates:
            return templates[template_choice]
        else:
            return print("Invalid Input")
    
    #
    while True:
        print(f"{BRIGHT_YELLOW}\n<--## Choose TEMPLATES ##-->{CEND}")
        print("1. ODCP")
        print("2. cWAN")
        print("3. CR2020")
        print("4. cWAN Black Box")
        print("\nSelect a template please..")
        
        template_choice = input("\nChoose a Template :")
        # if template_choice.strip() == "":
        #     template_data = get_custom_input(select_template(template_choice))
        if template_choice in ['1','2','3','4']:
            selected_template = select_template(template_choice)
        else:
            print(f"{RED}Invalid Input!{CEND}")




        if selected_template:
            print()
            for key, value in selected_template.items():
                print(value)
            
            print("\nWould you like to edit the template?")
            if input("\nDo you edit the template (y/n): ").lower() != 'n':
                get_custom_input(selected_template)
            
                # Example of breaking the loop or continuing based on some condition
                if input("\nDo you want to continue? (y/n): ").lower() != 'n':
                    break
            
            else:
                break



    # # Constants for positioning barcodes on the page
    # x_start = 10 * mm  # Starting x position (20mm from the left margin)
    # y_start = 235 * mm  # Starting y position (280mm from the bottom margin)
    # x_increment = 0  # No horizontal spacing between barcodes in a row
    # y_increment = 18 * mm  # Vertical spacing between barcodes (adjusted to make them closer)
    # barcode_width = 90 * mm  # Barcode width
    # barcode_height = 16 * mm  # Reduced barcode height




    #directory to store barcode, deleted when program done or when ctrl+c presses
    directory = "./bufferDEL"

    # Declearing comumn names as per excel sheet
    EXCEL_COLUMN_1_NAME = 'SN'
    EXCEL_COLUMN_2_NAME = 'IMEI'


    chosen_excel_file = chooseFile("./data/")

    serial_number_list, wan_mac_list = validateExcel(chosen_excel_file, EXCEL_COLUMN_1_NAME, EXCEL_COLUMN_2_NAME)

    no_of_barcode = len(serial_number_list)




    pdf_path = create_stickers(selected_template, serial_number_list)

    print()
    print(f"Sticker PDF created: {pdf_path}")


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



    # # Load data from Excel file
    # workbook = openpyxl.load_workbook('input.xlsx')
    # sheet = workbook.active

    # # Constants for positioning barcodes on the page
    # x_start = 10 * mm  # Starting x position (20mm from the left margin)
    # y_start = 235 * mm  # Starting y position (280mm from the bottom margin)
    # x_increment = 0  # No horizontal spacing between barcodes in a row
    # y_increment = 18 * mm  # Vertical spacing between barcodes (adjusted to make them closer)
    # barcode_width = 90 * mm  # Barcode width
    # barcode_height = 16 * mm  # Reduced barcode height

    location = chooseFile("./")
    # Load the Excel file into a pandas DataFrame
    # Make sure to replace 'your_excel_file.xlsx' with the actual path to your Excel file
    df = pd.read_excel(location, engine='openpyxl')  # Ensure you have 'openpyxl' installed for .xlsx files

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



    print()
    pdf_path = create_stickers(data1, data2, data3, barcodes)

    print()
    print(f"Sticker PDF created: {pdf_path}")




    #'module_width': 10, 'module_height': 80, "font_size": 20*5, "text_distance": 28


def cartonStickers():

    # Function to print the data in a formatted way
    # def print_formatted_data(box_data):
        # for box, details in box_data.items():
        #     print(f"Data for {box}:")
        #     print(f"{'RSN':<20} {'MAC':<20}")
        #     # Skip the first element of each list because it's the header based on how we simulated the data
        #     for sn, mac in zip(details['RSN'][1:], details['MAC'][1:]):
        #         print(f"{sn:<20} {mac:<20}")
        #     print("\n")

    # Function to process the Excel data and extract RSN and MAC for each box
    def extract_sn_mac(excel_data):
        # Dictionary to hold the box data
        box_data = {}
        boxvalue4search = ['box', 'box no']  # List of strings to search for
        current_box = None  # To keep track of the current box number

        for index, row in excel_data.iterrows():
            # Check if any of the terms in boxvalue4search are in the first cell of the row
            if any(search_term in str(row[0]).lower().replace('.','') for search_term in boxvalue4search):
                current_box = str(row[0])  # Update the current box number
                box_data[current_box] = {'RSN': [], 'MAC': []}
            elif pd.notnull(row[1]) and pd.notnull(row[2]) and current_box:
                box_data[current_box]['RSN'].append(row[1])
                box_data[current_box]['MAC'].append(row[2])

        return box_data

    # Function to extract data below the search values in the same column
    # def extract_data_below_values(sheet, search_values):
    #     # Iterate through all cells and search for the specified values
    #     for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
    #         for cell in row:
    #             # Normalize the cell value for comparison
    #             normalized_cell_value = str(cell.value).lower().replace(' ', '').replace('.', '').replace('\xa0', '')

    #             # Search for the normalized values
    #             if any(search_value in normalized_cell_value for search_value in search_values):
    #                 # Get the row and column indices
    #                 row_index = cell.row
    #                 col_index = cell.column

    #                 # Extract the data below the values in the same column
    #                 data = [sheet.cell(i, col_index).value for i in range(row_index + 1, sheet.max_row + 1)]

    #                 # Return a tuple with both data and column index
    #                 return pd.DataFrame({f"{sheet.title}_{col_index}": data}), col_index

    #     # If none of the search values are found, extract all non-empty values from the first column
    #     first_column_data = [sheet.cell(i, 1).value for i in range(1, sheet.max_row + 1) if sheet.cell(i, 1).value is not None]

    #     # Return a DataFrame with the first column data and None for the column index
    #     return pd.DataFrame({f"{sheet.title}_1": first_column_data}), None

    def boxStickers(data, ctnno, total_boxes):

        start_time = time.time()
        
        # MODEL MSN
        # msn = "M5005491008BKA00" #Add without box number
        msn = "M5005571808BKA00"
            
        ean = "0796554198316"

        #ctnno = input("Enter Carton No. : ")

        # Create PDF canvas
        pdf = canvas.Canvas(f'Cartion_Box_Sticker_pdf/Box{ctnno}.pdf', pagesize=A4)




        #collecting data fomr excel file and storing it in an array
        val1 = []
        val2 = []

        # Replace values
        RSN_columnName = 'RSN' 
        MAC_columnName = 'MAC'

        val1 = data['RSN'][1:]
        val2 = data['MAC'][1:]

        no_of_barcode = len(val1)

        # Barcode per box
        barcode_per_page = 12 #int(input("Enter Barcode Per Page (12 max) :"))


        
        
        

        if int(ctnno) <= 9:
            ctnno = "0"+str(ctnno)

        msn = (msn+ str(ctnno))

        #no = int(input('Start RSN for this sheet : '))


        
        # Variables to track page count and barcode count
        page_count = 0
        barcode_count = 0

        # Iterate through rows in the Excel sheet
        i = 0

        font_size = 10

        # Constants for positioning barcodes on the page
        x_start = 10 * mm  # Starting x position (20mm from the left margin)
        y_start = 235 * mm  # Starting y position (280mm from the bottom margin)
        x_increment = 0  # No horizontal spacing between barcodes in a row
        y_increment = 18 * mm  # Vertical spacing between barcodes (adjusted to make them closer)
        barcode_width = 90 * mm  # Barcode width
        barcode_height = 16 * mm  # Reduced barcode height

        #wight of one in kg
        oneBox_Gross_Weight = 1.241
        oneBox_Net_Weight = 1.016



        while i < len(val1): 

            # Calculate position on the page
            x = x_start+10 # Adjust space from right
            y = y_start - ((barcode_count % barcode_per_page) * y_increment) - 50  # Adjusted to place 14 barcodes in a column

            x1 = x_start+290 # Adjust space from right
            y1 = y_start - ((barcode_count % barcode_per_page) * y_increment) - 50

            x2 = 20
            y2 = 800
            
            ## Above text left ##
            pdf.setFont("Helvetica", 15)

            pdf.drawString(x2, y2, f"Commodity: Credo CR-3120-OD")
            pdf.drawString(x2, y2-20, f"Color: White")
                                # MODEL PO I03/45005571
            pdf.drawString(x2, y2-40, f"PO: I03/450055718")
            
            pdf.drawString(x2, y2-60, f"Date:02/2024")

            pdf.drawString(x2, y2-100, f"Gross Wt : {round(oneBox_Gross_Weight * no_of_barcode, 2)} Kg")
            pdf.drawString(x2, y2-120, f"Net Wt. : {round(oneBox_Net_Weight * no_of_barcode, 2)} Kg")

            ## Above text right ##
            pdf.drawString(x2+400, y2, f"Carton No. : {ctnno} of {total_boxes}") #make dynamic value
            pdf.drawString(x2+480, y2-20, f"Qty : {no_of_barcode}")


            msn1 = Code128(msn, writer=ImageWriter())
            msn_image = msn1.render(writer_options={'module_width': 4, 'module_height': 80, "font_size": 20*5, "text_distance": 40, "quite_zone": 10})
            msn_image_filename = (f"./bufferDEL/msn_barcode.png")
            msn_image.save(msn_image_filename)
            pdf.drawImage(msn_image_filename, x2+310, y2-75, width=barcode_width, height=barcode_height)
            pdf.drawString(x2+260, y2-50, f"MSN:")


            ean1 = Code128(ean, writer=ImageWriter())
            ean_image = ean1.render(writer_options={'module_width': 4, 'module_height': 80, "font_size": 20*5, "text_distance": 40, "quite_zone": 10})
            ean_image_filename = (f"./bufferDEL/ean_barcode.png")
            ean_image.save(ean_image_filename)
            pdf.drawImage(ean_image_filename, x2+310, y2-125, width=barcode_width, height=barcode_height)
            pdf.drawString(x2+260, y2-100, f"EAN:")



            pdf.setFont("Helvetica-Bold", font_size-1)
            
            ## RSN placed on the left side of the page. 

            # Insert barcode image into PDF
            rsn = Code128(val1[i], writer=ImageWriter())
            rsn_image = rsn.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*5  , "text_distance": 40, "quite_zone": 10})
            rsn_image_filename = (f"./bufferDEL/rsn_barcode_{i}.png")
            rsn_image.save(rsn_image_filename)
            pdf.drawImage(rsn_image_filename, x, y-18, width=barcode_width, height=barcode_height+12)
            pdf.drawString(x-35, y+25, f"RSN:")
            #pdf.drawString(x-30, y+15, f"{no}")


            macid = Code128(val2[i], writer=ImageWriter())
            macid_image = macid.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*5, "text_distance": 40, "quite_zone": 10})
            macid_image_filename = (f"./bufferDEL/macid_barcode{i}.png")
            macid_image.save(macid_image_filename)
            pdf.drawImage(macid_image_filename, x1, y1-18, width=barcode_width, height=barcode_height+12)   
            pdf.drawString(x2+311, y2-130, f"MAC ID: ")
            
            







            # Remove the barcode image file
            os.remove(rsn_image_filename)
            os.remove(macid_image_filename)

            # Increment the barcode count
            barcode_count += 1

            # Check if a new page is needed
            if barcode_count == barcode_per_page:
                # Reset parameters for new page
                page_count += 1
                barcode_count = 0

                # Show the current page
                pdf.showPage()




            print_progress_bar(i+1, start_time, no_of_barcode)
            #no = no+1
            i = i + 1 
        # Save the PDF
        print()
        print(f"File save at : Cartion_Box_Sticker_pdf/Box{ctnno}.pdf")
        pdf.save()

    def readEXCELnValidate():

        df = pd.read_excel(location)  # You would use the actual path to your Excel file

        # Extract the data for all boxes once
        extracted_data = extract_sn_mac(df)

        # Extract box numbers and convert them to integers for proper numeric sorting
        box_numbers = [int(''.join(filter(str.isdigit, box))) for box in extracted_data.keys()]

        # Now sort the box numbers in numeric order
        sorted_box_numbers = sorted(box_numbers)
        total_boxes = max(sorted_box_numbers)

        try:
            for box_number in sorted_box_numbers:
                # Convert back to the original box format if needed, or directly use box_number if applicable
                box_key = f"{searchValue}{box_number}"  # Adjust format as necessary based on how your keys are structured
                if box_number >= startFrom:
                    print(f"{searchValue}{box_number}")
                    boxStickers(extracted_data[f"{searchValue}{box_number}"], str(box_number), total_boxes)
        except KeyError as ke:
            print()
            print(f"{RED}Data not formatted properly. Please format data as below :-{CEND}")
            print(f"\n{RED}Value mismatch at {ke} {CEND}")
            print()
            exampleData = f"""{YELLOW}EXAMPLE DATA:-

                    |******> For BOX 1 <******|

            +-------+-----------------+-------------------+
            | BOX 1 |                 |                   |
            +-------+-----------------+-------------------+
            |  No.  | RSN             | MAC               |
            +-------+-----------------+-------------------+
            |   1   | RCRODBK01290308 | 44:B5:9C:00:48:C8 |
            |   2   | RCRODBK01290336 | 44:B5:9C:00:49:00 |
            |   3   | RCRODBK01290334 | 44:B5:9C:00:48:FC |
            |   4   | RCRODBK01290328 | 44:B5:9C:00:48:F0 |
            |   5   | RCRODBK01290312 | 44:B5:9C:00:48:D0 |
            +-------+-----------------+-------------------+{CEND}"""

            print(exampleData)   
            print()


    # Load the entire workbook once, instead of in the loop
    location = chooseFile("./boxData/")
    # MODEL DATASET
    # location = 'MODEL_carton_jio300_7feb24.xlsx'

    print()
        

    startFrom = int(input("Enter Box no. to start with : "))  # Assuming we are starting from box 1 for the sake of demonstration

    searchValue = 'box'

    readEXCELnValidate()







def userInterface():

    try:
        while True:

            
            print()
            print("Enter 1 to create Router Body Stickers.")
            print("Enter 2 to create Router BOX Stickers.")
            print("Enter 3 to create Router Carton Stickers.")
            print()
            print("0 to EXIT")
            print()
            choice = int(input("Enter Choice : "))
            print()
            if choice == 1:
                router_body_stickers()
                continue

            if choice == 2:
                router_box_stickers()
                continue

            if choice == 3:
                cartonStickers()
                continue

            if choice == 0:
                break
        
    except KeyboardInterrupt:
            print()
            print_banner("Rage quite initiated!! Bye!")
        
    # except Exception as e:
    #         print("Falling apart with error :-\n")
    #         print(e)

userInterface()