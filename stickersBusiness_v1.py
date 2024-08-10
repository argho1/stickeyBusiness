import json
import os
import re
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

        if not files:
            print(f"\n{RED}Folder {folder_path} is EMPTY!! Please insert Your Excel files in that folder.{CEND}\n")
            sys.exit(1)
        
        else:
            print()
            print_banner(f"Available Excel files in {folder_path}:")

            for i, file in enumerate(files, 1):
                file_path = os.path.join(folder_path, file)
                # Get the modification timestamp and format it as a date
                modified_date = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                print(f"{i}. {file} - Last Modified: {modified_date})")
            
            print("\n")
            # Select file
        
            chosen_file = choose_file(files)
            print()
            print(f"\033[32mYou chose: {chosen_file}\033[0m\n")
            #chosen_file_location = f"{folder_path}{chosen_file}"
            return chosen_file #_location

    except KeyboardInterrupt:
        print("\nCTRL + C pressed, getting the duck outta here.\n")
        sys.exit(1)

def delete_contents_of_directory(directory):
    if not os.path.exists(directory):
        print(f"The directory {directory} does not exist.")
        return

    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.remove(item_path)
        
        except Exception as e:
            print(f"Failed to delete {item_path}. Reason: {e}")

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
    sys.stdout.write(f'\rPage {page}/{total_pages} |[{bar}| {percentage:.2f}% Done. Elapsed: {elapsed_time_formatted}, Remaining: {remaining_time_formatted}')
    sys.stdout.flush()


def check_and_create_file():
    print()
    directory_list = [
        './bufferDEL',
        './Router_Body_Stickers_PDF', 
        './Router_BOX_Stickers_PDF', 
        './Router_Carton_Stickers_PDF',
        './templates_JSON',
        './ExcelData',
        './ModelExcelData',
        ]

    for directory in directory_list:
        # Check if the directory exists
        if not os.path.exists(directory):
            try:
                # Create the directory if it does not exist
                os.makedirs(directory)
                print(f"{BRIGHT_BLUE}Missing directory created: {directory}{CEND}")
            except Exception as e:
                print(f"{RED}An error occurred while creating directory {directory}: {e}{CEND}")

        if directory == './templates_JSON':

            BODY_template = {
                "1": {
                    "commodity_text": "Outdoor Router",
                    "model_text": "Credo CR-3120-OD",
                    "input_text": "48V PoE",
                    "eanno": "0796554198316"
                    },

                "2": {
                    "commodity_text": "CWAN",
                    "model_text": "CR3181-X",
                    "input_text": "240V AC"
                    },

                "3": {
                    "title"   : "Cellular Router",
                    "model"   : "Model : CR2020",
                    "power"   : "Power : 9-36V / 1.5A",
                    "version" : "Version : V2A-S230E",
                    "bands1"  : "Bands : LTE FDD(B1/B3/B5/B8)",
                    "bands2"  : "       LTE TDD(B34/B38/B39/B40/B41)"
                },
                
                "4": {
                    "title"   : "cWAN",
                    "model"   : "Model : # FROM EXCEL #",
                    "power"   : "Power : 220V AC",
                    "imei1"   : "IMEI1 : # FROM EXCEL #",
                    "imei2"   : "IMEI2 : # FROM EXCEL #"
                }
            }

            BOX_template = {

                "data1": "Commodity",
                "data1a": "Credo CR-3120-OD Router",
                "data2": "Manufactured By",
                "data2a": "Tenet Networks Private Limited",
                "data3": "Net Quantity",
                "data3a": "1 Outdoor Router + 1 Patch Cord",
                "data31": "1 POE Adapter + 1 clamp",
                "data4": "Month & Year of Manufacture: 02/2024",
                "data5": "Office Address",
                "data5a": "A-541, Logix Technova Sector-132",
                "data51": "Noida-201305 U.P. India",
                "data6": "Customer Care No.",
                "data6a": "+91 120-4165905",
                "data7": "Email ID",
                "data7a": "info@tenetnetworks.com"
            }

            with open(f"{directory}\\BODY_template.json", 'w') as json_file:
                json.dump(BODY_template, json_file, indent=4)

            with open(f"{directory}\\BOX_template.json", 'w') as json_file:
                json.dump(BOX_template, json_file, indent=4)
            
        if directory == './ModelExcelData':

            data = {
                "ODCP_BODY_MODEL_DATA.xlsx" : {

                    "SN": ["RCRODBK01290301", "RCRODBK01290302", "RCRODBK01290303", "RCRODBK01290304"],
                    "WAN_MAC":["44B59C0048BA", "44B59C0048BC", "44B59C0048BE", "44B59C0048C0"],
                },

                "cWAN_BlackBox_BODY_MODEL_DATA.xlsx" : {

                    "SN":       ["CRARM311736E6D", "CRARM311736E7D", "CRARM311736E8D"],
                    "IMEI1":    [ "861942058180001",  "861942058180002",  "0"],
                    "IMEI2":    ["860965062570011", "0", "0"],
                    "MODEL":    ["CR1211-A", "CR1111-A", "CR1011-A"]
                },

                "BOX_MODEL_DATA.xlsx": {
                    "BOX 1": {
                        "RSN": [
                            "RCRODBK01290308", "RCRODBK01290336", "RCRODBK01290334", "RCRODBK01290328",
                            "RCRODBK01290312", "RCRODBK01290494", "RCRODBK01290495", "RCRODBK01290320",
                            "RCRODBK01290307", "RCRODBK01290323", "RCRODBK01290497", "RCRODBK01290304",
                            "RCRODBK01290326", "RCRODBK01290332", "RCRODBK01290317", "RCRODBK01290318",
                            "RCRODBK01290316", "RCRODBK01290305"
                        ],
                        "MAC": [
                            "44:B5:9C:00:48:C8", "44:B5:9C:00:49:00", "44:B5:9C:00:48:FC", "44:B5:9C:00:48:F0",
                            "44:B5:9C:00:48:D0", "44:B5:9C:00:4A:3C", "44:B5:9C:00:4A:3E", "44:B5:9C:00:48:E0",
                            "44:B5:9C:00:48:C6", "44:B5:9C:00:48:E6", "44:B5:9C:00:4A:42", "44:B5:9C:00:48:C0",
                            "44:B5:9C:00:48:EC", "44:B5:9C:00:48:F8", "44:B5:9C:00:48:DA", "44:B5:9C:00:48:DC",
                            "44:B5:9C:00:48:D8", "44:B5:9C:00:48:C2"
                        ]
                    },
                    "BOX 2": {
                        "RSN": [
                            "RCRODBK01290315", "RCRODBK01290306", "RCRODBK01290367", "RCRODBK01290343",
                            "RCRODBK01290360", "RCRODBK01290428", "RCRODBK01290416", "RCRODBK01290363",
                            "RCRODBK01290351", "RCRODBK01290319", "RCRODBK01290329", "RCRODBK01290302",
                            "RCRODBK01290325", "RCRODBK01290420", "RCRODBK01290331", "RCRODBK01290392",
                            "RCRODBK01290350", "RCRODBK01290418"
                        ],
                        "MAC": [
                            "44:B5:9C:00:48:D6", "44:B5:9C:00:48:C4", "44:B5:9C:00:49:3E", "44:B5:9C:00:4A:C6",
                            "44:B5:9C:00:49:30", "44:B5:9C:00:49:B8", "44:B5:9C:00:49:A0", "44:B5:9C:00:49:36",
                            "44:B5:9C:00:49:1E", "44:B5:9C:00:48:DE", "44:B5:9C:00:48:F2", "44:B5:9C:00:48:BC",
                            "44:B5:9C:00:48:EA", "44:B5:9C:00:49:A8", "44:B5:9C:00:48:F6", "44:B5:9C:00:49:70",
                            "44:B5:9C:00:49:1C", "44:B5:9C:00:49:A4"
                        ]
                    },
                    "BOX 3": {
                        "RSN": [
                            "RCRODBK01290335", "RCRODBK01290346", "RCRODBK01290301", "RCRODBK01290341",
                            "RCRODBK01290311", "RCRODBK01290498", "RCRODBK01290359", "RCRODBK01290429",
                            "RCRODBK01290309", "RCRODBK01290314", "RCRODBK01290357", "RCRODBK01290310",
                            "RCRODBK01290348", "RCRODBK01290349", "RCRODBK01290324", "RCRODBK01290330",
                            "RCRODBK01290496", "RCRODBK01290327"
                        ],
                        "MAC": [
                            "44:B5:9C:00:48:FE", "44:B5:9C:00:49:14", "44:B5:9C:00:48:BA", "44:B5:9C:00:49:0A",
                            "44:B5:9C:00:48:CE", "44:B5:9C:00:4A:44", "44:B5:9C:00:49:2E", "44:B5:9C:00:49:BA",
                            "44:B5:9C:00:48:CA", "44:B5:9C:00:48:D4", "44:B5:9C:00:49:2A", "44:B5:9C:00:48:CC",
                            "44:B5:9C:00:49:18", "44:B5:9C:00:49:1A", "44:B5:9C:00:48:E8", "44:B5:9C:00:48:F4",
                            "44:B5:9C:00:4A:40", "44:B5:9C:00:48:EE"
                        ]
                    },
                },
            }

            blank_rows = 3
            df_list = []
            for key, value in data.items():
                file_path = f"{directory}\\{key}"
                if key == 'BOX_MODEL_DATA.xlsx':
                    # Iterate through the JSON data
                    for section_name, section_data in value.items():
                        # Convert each section to a DataFrame
                        df = pd.DataFrame(section_data)

                        # shift 1 column right
                        df.insert(0, None, '')

                        df.columns = [None] * len(df.columns)
                        
                        box_number = pd.DataFrame([[section_name,'','']], index=["Header"], columns=df.columns)

                        headers = pd.DataFrame([['','RSN', 'MAC']], columns=df.columns)

                        # Append the blank rows DataFrame to the list
                        df_list.append(box_number)

                        df_list.append(headers)

                        # Append the section DataFrame to the list
                        df_list.append(df)

                        # Create a DataFrame of blank rows with the same columns
                        blanks = pd.DataFrame("", index=range(blank_rows), columns=df.columns)
                        
                        # Append the blank rows DataFrame to the list
                        df_list.append(blanks)

                    # Concatenate all DataFrames in the list into a single DataFrame
                    final_df = pd.concat(df_list, ignore_index=True)
                    # Print the final DataFrame
                    
                    final_df.to_excel(file_path, index=False)

                else:
                    df = pd.DataFrame(data[key])

                    df.to_excel(file_path, index=False)

def banner():
    print(
        f"""
     ______________________________________________________________
    /                                                              ==
{RED}
                                       |          
                                      ...
                                     .:+:-.
                                    ..::*=:..      {CEND} {GREEN}
                                ....::####:...
                            ...-=::--*##=-::::=.
                            ....::##############=. {CEND} {BRIGHT_BLUE}
                    ........::#@#############@:.
                    ..:%::=::::-#################%
                ....::*#####=##################+ {CEND} {YELLOW}
            ......:==%#########################. 
            :::#:===+##########################-:.
            .....:--%#########################:
                ...::*#####-##################+ {CEND} {BRIGHT_BLUE}
                    .:=::-::::=#################@
                    .........:#+#############+:.
                            ....::############%@+. {CEND} {GREEN}
                            ..:-:::-###=:::..:.
                                ....::###%:...
                                    ...:-=:..      {CEND} {RED}
                                    .:%::.
                                      ...
                                       |         {CEND} 
            Stickey Business
            - by Argho Sinha
    \______________________________________________________________==

"""
    )






def router_body_stickers():
    print()

    #code to make stickrs pdf
    def create_stickers(selected_template, dataset):

        serial_number_list = dataset['SN']
        wan_mac_list = dataset['WAN_MAC']

        # Check if any value in serial_number_list is 0
        has_zero_in_sn = any(sn == 0 for sn in serial_number_list)

        # Check if any value in wan_mac_list is 0
        has_zero_in_wan_mac = any(mac == 0 for mac in wan_mac_list)

        no_of_barcode = len(serial_number_list)

        if has_zero_in_sn:
            print()
            print("\x1b[31m ERROR : VALUES are MISSING in SN column!! \x1b[0m")
            print()
            print("\x1b[31mPlease Check Excel Sheet and Try Again!!\x1b[0m")
            print()
            sys.exit(1)
        if has_zero_in_wan_mac:
            print()
            print("\x1b[31mERROR : VALUES are MISSING in WAN_MAC column!! \x1b[0m")
            print()
            print("\x1b[31mPlease Check Excel Sheet and Try Again!!\x1b[0m")
            print()
            sys.exit(1)



        commodity_text_print = selected_template['commodity_text']
        model_text_print = selected_template['model_text']
        input_text_print = selected_template['input_text']

        try:
            start_time = time.time()

            sticker_name = input("Please Enter sticker name : ")
            print()

            sticker_pdf_name = f"{DOWNLOAD_DIR}{sticker_name}_body_stickers.pdf"

            # Create a new PDF document
            c = canvas.Canvas(sticker_pdf_name, pagesize=A4)

            # Set the size and position of the stickers
            sticker_width = 200
            sticker_height = 180
            margin = 50
            page_width, page_height = A4
            num_stickers_per_row = 2
            num_stickers_per_col = 3
            stickers_per_page = num_stickers_per_row * num_stickers_per_col
            num_pages = (no_of_barcode + stickers_per_page - 1) // stickers_per_page
            i = 1

            for page in range(num_pages):
                start_idx = page * stickers_per_page
                end_idx = min((page + 1) * stickers_per_page, no_of_barcode)

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
                    c.drawString(text_x + 18, text_y - 16, f"Commodity : {commodity_text_print}")
                    c.drawString(text_x + 18, text_y - 28, f"Model : {model_text_print}")
                    c.drawString(text_x + 18, text_y - 40, f"Input : {input_text_print}")

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
                    c.drawString(barcode_x-45, barcode_y-5, 'SN:')
                    # c.drawString(barcode_x-44, barcode_y-10, f'{i}')
                    
                    macid = Code128(str(wan_mac_list[barcode_idx]), writer=ImageWriter())
                    macid_image = macid.render(writer_options={'module_width': 2, 'module_height': 80, "font_size": 30*2, "text_distance": 25, "quite_zone": 10})
                    macid_image_filename = f"./bufferDEL/WanMac_barcode_{wan_mac_list[barcode_idx]}.png"
                    macid_image.save(macid_image_filename)
                    # barcode_x for - to move left barecode_y to - to move down
                    c.drawImage(macid_image_filename, barcode_x-25, barcode_y-85, width=150, height=45)
                    c.drawString(barcode_x-45, barcode_y-57, f'MAC:')

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

            os.startfile(sticker_pdf_name)
            delete_contents_of_directory(BUFFER_DIR)
            return sticker_pdf_name
        
        
        except KeyboardInterrupt:
            print()
            print_banner("Rage Quit Inititated!!")
            delete_contents_of_directory(BUFFER_DIR)
            sys.exit(1)


    def margined_body_sticker(context, selected_template, dataset):

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
        
        def draw_sticker(c, x, y, width, height, sn, imei1=None, imei2=None, model_number=None):
            
            # STICKER FRAME
            c.line(x, y, x, y + 40 * mm + 3.5 * mm)  # Left border
            
            c.line(x, y, x + 49 * mm + 5.5 * mm, y)  # Button border
            
            c.line(x, y + 40 * mm + 3.5 * mm, x + 42 * mm + 5 * mm, y + 40 * mm + 3.5 * mm)  # Top border
            
            c.line(x + 49 * mm + 5.5 * mm, y , x + 49 * mm + 5.5 * mm, y + 33 * mm + 2.5 * mm)  # Right border
            
            c.line(x + 49 * mm + 5.6 * mm, y + 33 * mm + 2.3 * mm, x + 42 * mm + 5 * mm, y + 40 * mm + 3.5 * mm)  # Right corner tilted border

            #BARCODE
            draw_barcode(c, x + 1, y + 83, sn, -90)

            
            
            modified_template_data = {}
            for key, value in selected_template.items():
                if key == 'imei1':
                    if imei1 != 0:
                        modified_template_data[key] = f"IMEI 1 : {imei1}"
                    else:
                        modified_template_data[key] = 0
                elif key == 'imei2':
                    if imei2 != 0:
                        modified_template_data[key] = f"IMEI 2 : {imei2}"
                    else:
                        modified_template_data[key] = 0
                elif key == 'model':
                    if context == "SN_IMEI_MODEL_TEMPLATE":
                        modified_template_data[key] = f"Model : {model_number}"
                    else:
                        pass
                else:
                    modified_template_data[key] = value


            if "CR12" in modified_template_data["model"]:
                font = 3
                align_x = x + 14 * mm # 14 default
                align_y = y + 27.5 * mm # 26 default
                space = 0

            elif "CR11" in modified_template_data["model"]:
                font = 3
                align_x = x + 14 * mm # 14 default
                align_y = y + 26 * mm # 26 default
                space = 0
            
            elif "CR10" in modified_template_data["model"]:
                font = 3
                align_x = x + 14 * mm # 14 default
                align_y = y + 24.5 * mm # 26 default
                space = 0
            
            elif len(modified_template_data) == 6 :
                font = 3
                align_x = x + 14 * mm # 14 default
                align_y = y + 30 * mm # 26 default
                space = 0
            
            else:
                print(f"{len(modified_template_data)} Wrong size dictionary passed, improve code, call Argho lol!")

            #TOP LINE
            c.line(align_x - 1.5 * mm, align_y + 5 * mm, x + width - 10, align_y + 5 * mm )  # Top border

            #HEAD
            text1="CREDO "
            text2="NETWORKS"
            head_font_size = font + 1.5
            draw_text(c, align_x, align_y + 7 * mm, text1, head_font_size, font='Helvetica-Bold')
            draw_text(c, align_x + 12 * mm, align_y + 7 * mm, text2, head_font_size)

            
            for key, value in modified_template_data.items():
                if key == 'title':
                    draw_text(c, align_x, space + align_y, f"{value}", font + 2 , font='Helvetica-Bold')
                    space = space - 10
                elif key in ['bands1','bands2']:
                    draw_text(c, align_x + 2, space + align_y - 1.55 * mm, f"{value}", font, font='Helvetica-Bold')
                    space = space - 10
                elif "imei" in key and modified_template_data[key] != 0:
                    draw_text(c, align_x + 2, space + align_y - 1.55 * mm , f"{value}", font + 1)
                    space = space - 10
                elif modified_template_data[key] != 0:
                    draw_text(c, align_x + 2, space + align_y - 1.55 * mm , f"{value}", font + 1)
                    space = space - 10

            #BOTTOM LINE
            c.line(align_x - 1.5 * mm , align_y + space , x + width - 10, align_y + space )  # Top border
            space = space - 10

            text="CREDO NETWORKS"
            #TAIL
            draw_text(c,align_x + 1 * mm , align_y + space , text, font + 1                                                                                                                                                                                                        , font='Helvetica-Bold')
            
        





        serial_number_list = dataset['SN']    

        if context == "SN_IMEI_MODEL_TEMPLATE":
            imei1_list = dataset['IMEI1']
            imei2_list = dataset['IMEI2']
            model_list = dataset['MODEL']
        
            

        
        try:
            sticker_name = input("Please Enter sticker name : ")

            sticker_pdf_name = f"{DOWNLOAD_DIR}{sticker_name}_stickers.pdf"

            # Create a PDF for output
            c = canvas.Canvas(sticker_pdf_name, pagesize=A4)
            width, height = A4  # width and height of the page

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
            for i in range(len(serial_number_list)):
                start_time = time.time()
                sn = serial_number_list[i]
                
                if context == "SN_IMEI_MODEL_TEMPLATE":
                    imei1 = int(imei1_list[i])
                    model_number = model_list[i]
                    
                    try :
                        imei2 = int(imei2_list[i])
                    except Exception:
                        imei2 = 0
                        pass

                
                # print(sn, " ", imei)
                row_num = (i // stickers_per_row) % rows_per_page
                column = i % stickers_per_row

                x = start_x + column * (sticker_width + margin + additional_space)
                y = start_y - row_num * (sticker_height + 15 * mm)
                
                if context == "SN_IMEI_MODEL_TEMPLATE":
                    draw_sticker(c, x , y, sticker_width, sticker_height, sn, imei1, imei2, model_number)
                else:
                    draw_sticker(c, x , y, sticker_width, sticker_height, sn)

                # Check if we need a new page
                if (i + 1) % (stickers_per_row * rows_per_page) == 0:
                    c.showPage()
                    start_y = page_height - margin - sticker_height  # Reset y position
                print_progress_bar(i+1, start_time, len(serial_number_list))


            # A4 dimensions in points
            width, height = A4

            c.save()

            os.startfile(sticker_pdf_name)

            print(f"\n{GREEN}PDF file created at {sticker_pdf_name}\n")
            
            delete_contents_of_directory(BUFFER_DIR)

        except KeyboardInterrupt:
            print()
            print_banner("Rage Quit Inititated!!")
            delete_contents_of_directory(BUFFER_DIR)
            sys.exit(1)





    def validate_N_list_Excel(chosen_template, chosen_excel_file, excel_column_names_list):
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(chosen_excel_file, engine='openpyxl')

        if chosen_template == "SN_MAC_TEMPLATE":
            
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

        elif chosen_template == "SN_IMEI_MODEL_TEMPLATE":
            
            exampleData="""\033[33m
                EXAMPLE DATA:-
                
                +----------------+-----------------+-----------------+----------+
                | SN             | IMEI1           | IMEI2           | MODEL    |
                +----------------+-----------------+-----------------+----------+
                | CRARM311736E6D | 861942058188336 | 860965062571024 | CR1211-A |
                | CRARM311736E7D | 861942058188337 |                 | CR1111-A |
                | CRARM311736E8D |                 |                 | CR1011-A |
                | CRARM311736E9D | 861942058188338 | 860965062571025 | CR1211-A |
                +----------------+-----------------+-----------------+----------+\033[0m"""
            
        elif chosen_template == "SN_TEMPLATE":
            
            exampleData="""\033[33m
                EXAMPLE DATA:-
                
                +----------------+
                | SN             |
                +----------------+
                | CRARM311736E6D |
                | CRARM311736E7D |
                | CRARM311736E8D |
                | CRARM311736E9D |
                +----------------+\033[0m"""



        try:
            column_data = {}
            # Select the columns by name
            for column in excel_column_names_list:
                # Replace empty values with '0' and convert to list
                column_data[column] = df[column].fillna(0).tolist()
            return column_data
            
        except KeyError as ke:
            print("\033[31mColumns are not properly named.\033[0m")
            print()
            print("\033[33mPlease have data in .xlsx format with column names like in the example below.\033[0m")
            print(exampleData)
            print(f"\n{RED}Opening {chosen_excel_file}, please check data!!{CEND}\n")
            os.startfile(chosen_excel_file)
            sys.exit(1)



        

    def get_custom_input(selected_template):

        edited_template = {}
        for key, value in selected_template.items():
            
            if 'excel' not in value.strip().lower().replace(" ",""):
                print("\nDefault : ", str(value))
                userInput = input("\nEnter data or Hit enter to select default data :")
                
                if userInput == "":
                    edited_template[key] = value
                else:
                    edited_template[key] = userInput
            else:
                edited_template[key] = value
                print("Value taken from Excel : ",value)

        for key, value in edited_template.items():
            print(value)

        return edited_template 
    

    def select_template(template_choice):

        with open('templates_JSON\\BODY_template.json', 'r') as file:
            templates = json.load(file)

        if template_choice in templates:
            return templates[template_choice]
        else:
            return print("Invalid Input")
    
    #
    while True:
        print(f"{YELLOW}\n<--## Choose TEMPLATES ##-->{CEND}")
        print("1. ODCP")
        print("2. cWAN")
        print("3. CR2020")
        print("4. cWAN Black Box")
        print("\nSelect a template please..")
        
        template_choice = input("\nChoose a Template :")

        if template_choice in ['1','2','3','4']:
            selected_template = select_template(template_choice)
        else:
            print(f"\n{RED}Invalid Input!{CEND}")
            continue


        if selected_template:
            print()
            for key, value in selected_template.items():
                print("\t\t\t\t",value)
            
            print("\nWould you like to edit the template?")
            if input("\nDo you edit the template (y/n): ").lower() != 'n':
                selected_template = get_custom_input(selected_template)
                # Example of breaking the loop or continuing based on some condition
                if input("\nDo you want to continue? (y/n): ").lower() != 'n':
                    break
            
            else:
                break


    #directory to store barcode, deleted when program done or when ctrl+c presses
    BUFFER_DIR = "./bufferDEL"
    DOWNLOAD_DIR = "./Router_Body_Stickers_PDF\\"
    

    chosen_excel_file = chooseFile("./ExcelData\\")


    if template_choice in ['1','2']:

        context = "SN_MAC_TEMPLATE"

        excel_column_names_list = ['SN', 'WAN_MAC']
        dataset = validate_N_list_Excel(context, chosen_excel_file, excel_column_names_list)

        pdf_path = create_stickers(selected_template, dataset)
        print()
        print(f"\n{GREEN}Sticker PDF created: {pdf_path}{CEND}\n")

    # Cellular router
    elif template_choice == '3':

        context = "SN_TEMPLATE"

        excel_column_names_list = ['SN']
        dataset = validate_N_list_Excel(context, chosen_excel_file, excel_column_names_list)

        margined_body_sticker(context, selected_template, dataset)
    
    # cWAN
    elif template_choice == '4':

        context = "SN_IMEI_MODEL_TEMPLATE"

        excel_column_names_list = ['SN', 'IMEI1', 'IMEI2', 'MODEL']
        dataset = validate_N_list_Excel(context, chosen_excel_file, excel_column_names_list)

        margined_body_sticker(context, selected_template, dataset)




def router_box_stickers():



    #code to make stickrs pdf
    def create_stickers():

        start_time = time.time()

        chosenName = input("Please enter a name for pdf file : ")
        print()
        # Create a new PDF document
        sticker_pdf_name = f"./Router_BOX_Stickers_PDF\\{chosenName}_stickers.pdf"
        c = canvas.Canvas(sticker_pdf_name, pagesize=landscape(A4))


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
                c.drawString(text_x, text_y, data["data1"])
                c.drawString(text_x+150, text_y, data["data1a"])
                c.drawString(text_x, text_y - 20, data["data2"])
                c.drawString(text_x+150, text_y - 20, data["data2a"])
                c.drawString(text_x, text_y - 40, data["data3"])
                c.drawString(text_x+150, text_y - 40, data["data3a"])
                c.drawString(text_x+160, text_y - 60, data["data31"])
                c.drawString(text_x, text_y - 85, data["data4"])
                c.drawString(text_x, text_y - 120, data["data5"])
                c.drawString(text_x+150, text_y - 120, data["data5a"])
                c.drawString(text_x+160, text_y - 140, data["data51"])
                c.drawString(text_x, text_y - 160, data["data6"])
                c.drawString(text_x+150, text_y - 160, data["data6a"])
                c.drawString(text_x, text_y - 180, data["data7"])
                c.drawString(text_x+150, text_y - 180, data["data7a"])

                # Generate and add barcode below data3
                barcode_x = sticker_x * 10.8 + 50
                barcode_y = sticker_y + sticker_height - 80

                
                rcno = Code128(sn_list[barcode_idx], writer=ImageWriter())
                rcno_image = rcno.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*4, "text_distance": 30, "quite_zone": 10})
                rcno_image_filename = f"./{BUFFER_DIR}/barcode_{page}_{idx}.png"
                rcno_image.save(rcno_image_filename)
                #barcode_x for - to move left barecode_y to - to move down
                c.drawImage(rcno_image_filename, barcode_x-25, barcode_y-10, width=150+70, height=55)
                c.drawString(barcode_x-70, barcode_y+18, f"RSN :")
                c.drawString(barcode_x-65, barcode_y+5, f"{i}")


                ean = Code128(eanno, writer=ImageWriter())
                ean_image = ean.render(writer_options={'module_width': 4, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
                ean_image_filename = f"./{BUFFER_DIR}/ean_barcode.png"
                ean_image.save(ean_image_filename)
                #barcode_x for - to move left barecode_y to - to move down
                c.drawImage(ean_image_filename, barcode_x-25, barcode_y-70, width=150+70, height=55)
                c.drawString(barcode_x-70, barcode_y-40, "EAN :")


                
                macid = Code128(wan_mac_list[barcode_idx], writer=ImageWriter())
                macid_image = macid.render(writer_options={'module_width': 2.8, 'module_height': 80, "font_size": 20*4, "text_distance": 30})
                macid_image_filename = f"{BUFFER_DIR}/barcode1_{page}_{idx}.png"
                macid_image.save(macid_image_filename)
                #barcode_x for - to move left barecode_y to - to move down
                c.drawImage(macid_image_filename, barcode_x-25, barcode_y-130, width=150+70, height=55)
                c.drawString(barcode_x-70, barcode_y-100, "MAC :")
                


                #ProgressBar
                print_progress_bar(i, start_time, no_of_barcode)
                i=i+1
            # Add a new page for the next set of stickers
            if page < num_pages - 1:
                c.showPage()

        # Save the PDF document
        c.save()
        
        os.startfile(sticker_pdf_name)
        delete_contents_of_directory(BUFFER_DIR)

        return sticker_pdf_name


    location = chooseFile("./ExcelData\\")
    BUFFER_DIR = "./bufferDEL"

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

        exampleData=f"""{BRIGHT_YELLOW}
    EXAMPLE DATA:-

    +----------------+-------------------+
    | SN             | WAN_MAC           |
    +----------------+-------------------+
    | RCRODBK01290001| 44:B5:9C:00:46:53 |
    | RCRODBK01290002| 44:B5:9C:00:46:55 |
    | RCRODBK01290003| 44:B5:9C:00:46:57 |
    | RCRODBK01290004| 44:B5:9C:00:46:59 |
    +----------------+-------------------+{CEND}"""



        print("Please have data in .xlsx format with comumn names as SN for Serial Number and WAN_MAC for WAN MAC like exaple below.")
        print(exampleData)
        os.startfile(location)
        sys.exit(1)

    no_of_barcode = len(sn_list)

    file_path = './templates_JSON/BOX_template.json'

    # IMPORT JSON Data
    with open(file_path, 'r') as file:
        data = json.load(file)



    # # Static Data
    # data1 = 'Commodity'
    # data1a =': Credo CR-3120-OD Router'
    # data2 = 'Manufactured By'
    # data2a =': Tenet Networks Private Limited'
    # data3 = 'Net Quantity'
    # data3a=': 1 Outdoor Router + 1 Patch Cord'
    # data31= ' + 1 POE Adapter + 1 clamp'
    # data4 = 'Month & Year of Manufacture: 02/2024'
    # data5 = 'Office Address'
    # data5a=': A-541, Logix Technova Sector-132'
    # data51= 'Noida-201305 U.P. India'
    # data6 = 'Customer Care No.'
    # data6a=': +91 120-4165905'
    # data7 = 'Email ID'
    # data7a=': info@tenetnetworks.com'
    barcodes = sn_list

    #Iteration Count
    #i = 1


    #ESN no.
    eanno =("0796554198316")
    print()



    print()
    pdf_path = create_stickers()

    print()
    print(f"\n{GREEN}Sticker PDF created at : {pdf_path}{CEND}")




def router_carton_stickers():

    # Function to process the Excel data and extract RSN and MAC for each box
    def extract_sn_mac(excel_data):
        # Dictionary to hold the box data
        box_data = {}
        boxvalue4search = ['box', 'boxno']  # List of strings to search for
        current_box = None  # To keep track of the current box number

        for index, row in excel_data.iterrows():
            # Check if any of the terms in boxvalue4search are in the first cell of the row
            if any(search_term in str(row[0]).lower().replace('.','').replace(" ","") for search_term in boxvalue4search):
                current_box = str(row[0])  # Update the current box number
                box_data[current_box] = {'RSN': [], 'MAC': []}
            elif pd.notnull(row[1]) and pd.notnull(row[2]) and current_box:
                box_data[current_box]['RSN'].append(row[1])
                box_data[current_box]['MAC'].append(row[2])

        return box_data

    def cartonStickers(data, ctnno, total_boxes, msn, ean):

        start_time = time.time()

        #ctnno = input("Enter Carton No. : ")

        # Create PDF canvas
        pdf = canvas.Canvas(f'.\Router_Carton_Stickers_PDF\\Box{ctnno}.pdf', pagesize=A4)

        #collecting data fomr excel file and storing it in an array
        val1 = []
        val2 = []

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
        pdf.save()
        print(f"{GREEN}File saved at : Router_Carton_Stickers_PDF/Box{ctnno}.pdf{CEND}\n")
        

    def readEXCELnValidate(df):

        # Extract the data for all boxes once
        extracted_data = extract_sn_mac(df)

        # Converting BOX1 to box1 for comparison. #lowercase #filter
        processed_data = {
            re.sub(r'[^a-z0-9]', '', key.lower()): value
            for key, value in extracted_data.items()
        }

        # Extract box numbers and convert them to integers for proper numeric sorting
        box_numbers = [int(''.join(filter(str.isdigit, box))) for box in processed_data.keys()]

        print(f"\n{YELLOW}No. of BOXES detected : {len(box_numbers)}{CEND}")

        # Now sort the box numbers in numeric order
        sorted_box_numbers = sorted(box_numbers)
        total_boxes = max(sorted_box_numbers)

        try:
            for box_number in sorted_box_numbers:
                # Convert back to the original box format if needed, or directly use box_number if applicable
                box_key = f"{searchValue}{box_number}"  # Adjust format as necessary based on how your keys are structured
                if box_number >= startFrom:
                    print(f"\n{BRIGHT_BLUE}Printing sticker for {searchValue}{box_number}{CEND}")
                    cartonStickers(processed_data[f"{searchValue}{box_number}"], str(box_number), total_boxes, msn, ean)
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
    location = chooseFile("./ExcelData/")
    # MODEL DATASET
    # location = 'MODEL_carton_jio300_7feb24.xlsx'

    print()

    msn = input("Enter MSN : ")
    ean = input("Enter EAN : ")    

    startFrom = int(input("Enter Box No. to start printing from : "))  # Assuming we are starting from box 1 for the sake of demonstration

    searchValue = 'box'

    df = pd.read_excel(location)  # You would use the actual path to your Excel file

    readEXCELnValidate(df)







def userInterface():

    try:
        while True:
            print()
            print(f"{YELLOW}+++++++++++++++++ MENU +++++++++++++++++{CEND}")
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
                router_carton_stickers()
                continue

            if choice == 0:
                break

            else:
                print(f"\n{RED}Invalid input!! Please try again..{CEND}\n")
                continue
        
    except KeyboardInterrupt:
            print()
            print_banner("Rage quite initiated!! Bye!")
        
    # except Exception as e:
    #         print("Falling apart with error :-\n")
    #         print(e)






# FUNCTION CALLING

banner()

check_and_create_file()

userInterface()