import win32print
import win32api
import tkinter as tk
from tkinter import simpledialog, messagebox

class PrintHelper:
    def __init__(self):
        self.printer_name = None
        self.setup_printer()

    def setup_printer(self):
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        self.printer_name = self.select_printer(printers)
        if self.printer_name:
            self.configure_printer()

    def select_printer(self, printers):
        root = tk.Tk()
        root.withdraw()
        printer_name = simpledialog.askstring("Select Printer", "Available Printers:\n" + "\n".join(printers))
        if printer_name and printer_name in printers:
            return printer_name
        else:
            messagebox.showerror("Error", "Invalid printer selected.")
            return None

    def configure_printer(self):
        try:
            hprinter = win32print.OpenPrinter(self.printer_name)
            print_info = win32print.GetPrinter(hprinter, 2)
            print_info['pDevMode'].Duplex = 1  # Example: Set duplex printing
            win32print.SetPrinter(hprinter, 2, print_info, 0)
            win32print.ClosePrinter(hprinter)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to configure printer: {e}")

    def print_document(self, file_path):
        if self.printer_name:
            try:
                win32api.ShellExecute(
                    0,
                    "print",
                    file_path,
                    f'/d:"{self.printer_name}"',
                    ".",
                    0
                )
                messagebox.showinfo("Success", "Document sent to printer.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to print document: {e}")
        else:
            messagebox.showwarning("Warning", "Printer is not set up.")

if __name__ == "__main__":
    helper = PrintHelper()
    document_path = simpledialog.askstring("Print Document", "Enter the full path of the document to print:")
    if document_path:
        helper.print_document(document_path)