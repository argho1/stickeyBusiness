import customtkinter as ctk


# Constants for mm to pixels conversion (assuming 96 DPI)
MM_TO_PX = 3.7795275591

# Convert mm to pixels function
def mm_to_px(mm):
    return mm * MM_TO_PX

# A4 size in pixels (assuming 96 DPI)
A4_WIDTH_PX = mm_to_px(210)
A4_HEIGHT_PX = mm_to_px(297)

# Create the main window
root = ctk.CTk()

# Create a Canvas widget
canvas = ctk.CTkCanvas(root, width=A4_WIDTH_PX, height=A4_HEIGHT_PX)
canvas.pack()

# Coordinates (example values)
x = 50
y = 50

# Draw the sticker frame
canvas.create_line(x, y, x, y + mm_to_px(40) + mm_to_px(3.5))  # Left border
canvas.create_line(x, y, x + mm_to_px(49) + mm_to_px(5.5), y)  # Bottom border
canvas.create_line(x, y + mm_to_px(40) + mm_to_px(3.5), x + mm_to_px(42) + mm_to_px(5), y + mm_to_px(40) + mm_to_px(3.5))  # Top border
canvas.create_line(x + mm_to_px(49) + mm_to_px(5.5), y, x + mm_to_px(49) + mm_to_px(5.5), y + mm_to_px(33) + mm_to_px(2.5))  # Right border
canvas.create_line(x + mm_to_px(49) + mm_to_px(5.6), y + mm_to_px(33) + mm_to_px(2.3), x + mm_to_px(42) + mm_to_px(5), y + mm_to_px(40) + mm_to_px(3.5))  # Right corner tilted border

# Function to draw a barcode (placeholder function)
def draw_barcode(canvas, x, y, sn, angle):
    # Draw a placeholder barcode (a simple line for illustration)
    canvas.create_line(x, y, x, y + 50, fill='black')

# Draw the barcode
draw_barcode(canvas, x + 1, y + 83, 'SN1234567890', -90)

# Run the application
root.mainloop()