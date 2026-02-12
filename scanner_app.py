
import tkinter as tk
from tkinter import messagebox, Label
import win32com.client
import os
import datetime
from PIL import Image, ImageTk

class ScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple Scanner")
        self.root.geometry("600x800")

        self.btn_scan = tk.Button(root, text="Scan Document", command=self.scan_document, font=("Arial", 16), height=2)
        self.btn_scan.pack(pady=20)

        self.status_label = tk.Label(root, text="Ready to scan", font=("Arial", 10))
        self.status_label.pack(pady=5)

        self.image_label = Label(root)
        self.image_label.pack(expand=True, fill="both", padx=10, pady=10)

        self.scanned_image = None

    def scan_document(self):
        try:
            self.status_label.config(text="Scanning... Please wait.")
            self.root.update()

            # Create WIA CommonDialog object
            # This will open the native Windows Scan dialog
            wia_dialog = win32com.client.Dispatch("WIA.CommonDialog")
            
            # ShowAcquireImage returns a WIA ImageFile object
            # 1 = ScannerDeviceType, 1 = IntentText (Unspecified/Color), 0 = Bias (Minimize size/Maximize quality)
            # We use default arguments to let the dialog handle it or just call ShowAcquireImage()
            image_file = wia_dialog.ShowAcquireImage()

            if image_file:
                # Save the image
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"scan_{timestamp}.jpg"
                file_path = os.path.abspath(filename)
                
                # WIA ImageFile SaveFile method
                if os.path.exists(file_path):
                    os.remove(file_path)
                image_file.SaveFile(file_path)

                self.status_label.config(text=f"Saved to: {file_path}")
                self.display_image(file_path)
            else:
                 self.status_label.config(text="Scan cancelled.")

        except Exception as e:
            # Check for cancellation (COM error)
            if "0x80210015" in str(e): # WIA_ERROR_USER_INTERVENTION usually means cancel
                 self.status_label.config(text="Scan cancelled by user.")
            else:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
                self.status_label.config(text="Error occurred.")
                print(e)
    
    def display_image(self, path):
        try:
            img = Image.open(path)
            
            # Resize image to fit in window
            window_width = self.root.winfo_width()
            window_height = self.root.winfo_height() - 150 # reserve space for buttons/labels
            
            if window_width > 10 and window_height > 10:
                img.thumbnail((window_width, window_height))

            self.scanned_image = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.scanned_image)
        except Exception as e:
            print(f"Failed to display image: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScannerApp(root)
    root.mainloop()
