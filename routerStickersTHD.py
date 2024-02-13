import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
import os
from concurrent.futures import ThreadPoolExecutor
import time
from tqdm import tqdm
import math
import os






# Updated function to generate both sn and WAN_MAC barcode images
def generate_barcode(row, page, idx, progress=None):
    try:
        # Generate sn Barcode
        sn = row['SN']  # Assuming 'SN' is the correct column name
        sn_image = Code128(sn, writer=ImageWriter()).render(writer_options={
            'module_width': 2.8,
            'module_height': 80,
            "font_size": 20*4,
            "text_distance": 30,
            "quiet_zone": 10
        })
        directory = "./bufferDEL"
        if not os.path.exists(directory):
            os.makedirs(directory)
        sn_image_filename = f"{directory}/SN_barcode_{page}_{idx}.png"
        sn_image.save(sn_image_filename)
        
        # Generate WAN_MAC Barcode
        wan_mac = row['WAN_MAC']
        macid_image = Code128(wan_mac, writer=ImageWriter()).render(writer_options={
            'module_width': 2,
            'module_height': 80,
            "font_size": 20*4,
            "text_distance": 30
        })
        macid_image_filename = f"{directory}/WanMac_barcode_{page}_{idx}.png"
        macid_image.save(macid_image_filename)
    finally:
        if progress is not None:
            progress.update(2)





# Function to determine the optimal number of threads
def optimal_thread_count():
    cpu_count = os.cpu_count() or 4  # Fallback to 4 if os.cpu_count() is None
    # Assuming the task is I/O-bound, you might experiment with this multiplier
    return max(4, cpu_count * 2)






# Function to read Excel and generate barcodes with progress bar and execution time measurement
def process_excel_and_generate_barcodes(excel_file, sheet_name='Sheet1'):
    start_time = time.time()
    
    # Read Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Setup progress bar - now considering we generate 2 barcodes per row
    progress = tqdm(total=len(df) * 2, desc='Generating Barcodes')
    
    # Generate barcodes in parallel
    with ThreadPoolExecutor(max_workers=optimal_thread_count()) as executor:
        futures = [executor.submit(generate_barcode, row, 0, idx, progress) for idx, row in df.iterrows()]
        for future in futures:
            future.result()  # Wait for all futures to complete
    
    progress.close()
    end_time = time.time()
    print(f"Completed in {end_time - start_time:.2f} seconds")






# Specify your Excel file path and sheet name
excel_file_path = 'jio300_7feb24.xlsx'
sheet_name = 'Sheet1'





process_excel_and_generate_barcodes(excel_file_path, sheet_name)
