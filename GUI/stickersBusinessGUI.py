import customtkinter
from PIL import Image, ImageTk
import barcode
from barcode.writer import ImageWriter
import os
import logging

# Set up logging
logging.basicConfig(level=logging.DEBUG)

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")

class BarcodeApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("500x500")
        self.root.title("Barcode Generator")

        self.main_frame = customtkinter.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.label = customtkinter.CTkLabel(self.main_frame, text="Enter value to convert to barcode:")
        self.label.pack(pady=10)

        self.entry = customtkinter.CTkEntry(self.main_frame)
        self.entry.pack(pady=10)

        self.generate_button = customtkinter.CTkButton(self.main_frame, text="Generate Barcode", command=self.generate_barcode)
        self.generate_button.pack(pady=10)

        myImage = customtkinter.CTkImage(light_image=Image.open("barcode.png"), size=(260,300))
        
        self.my_lable = customtkinter.CTkLabel(root, text="", image=myImage)
        self.my_lable.pack(pady=10)

        self.barcode_label = customtkinter.CTkLabel(self.main_frame, text="")
        self.barcode_label.pack(pady=20)

    def generate_barcode(self):
        data = self.entry.get()
        if not data:
            self.barcode_label.configure(text="Please enter a value.")
            return

        try:
            barcode_filename = self.create_barcode(data)
            if os.path.exists(barcode_filename):
                barcode_image = Image.open(barcode_filename)
                barcode_photo = ImageTk.PhotoImage(barcode_image)
                self.barcode_label.configure(image=barcode_photo, text="")
                self.barcode_label.image = barcode_photo  # Keep a reference to avoid garbage collection
            else:
                self.barcode_label.configure(text="Error generating barcode.")
        except Exception as e:
            logging.error(f"Error generating barcode: {e}")
            self.barcode_label.configure(text=f"Error: {e}")

    def create_barcode(self, data):
        try:
            CODE128 = barcode.get_barcode_class('code128')
            barcode_image = CODE128(data, writer=ImageWriter())
            filename = "barcode"
            barcode_image.save(filename)
            logging.debug(f"Barcode saved to {filename}")
            return filename
        except Exception as e:
            logging.error(f"Error creating barcode: {e}")
            raise

if __name__ == "__main__":
    root = customtkinter.CTk()
    app = BarcodeApp(root)
    root.mainloop()
