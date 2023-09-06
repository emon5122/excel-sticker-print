import openpyxl
import pyprinter
import win32print

def print_cartoon_stickers_from_excel(printer_name):
    try:
        workbook = openpyxl.load_workbook("sticker.xlsx")
        sheet = workbook["sticker"]
    except FileNotFoundError:
        print("File sticker.xlsx not found.")
        return
    except KeyError:
        print("Sheet sticker not found in the Excel file.")
        return

    try:
        printer = pyprinter.get_printer(printer_name)
    except pyprinter.PrinterNotFound:
        print(f"Printer '{printer_name}' not found.")
        return

    for row in sheet.iter_rows(min_row=2, values_only=True):
        cartoon_num, sticker_text = row
        sticker_text = f"Cartoon #{cartoon_num}: {sticker_text}"
        printer.print(sticker_text)
        print(f"Printed: {sticker_text}")

def list_available_printers():
    if printers := win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL, None, 1
    ):
        for i, printer_info in enumerate(printers, start=1):
            print(f"{i}. {printer_info[2]}")
        return printers
    else:
        print("No printers found.")
        return None

def main():
    print("Select a printer by entering its number:")
    if available_printers := list_available_printers():
        printer_number = input()
        try:
            printer_index = int(printer_number) - 1
            selected_printer_name = available_printers[printer_index][2]
            print_cartoon_stickers_from_excel(selected_printer_name)
        except (ValueError, IndexError):
            print("Invalid selection. Please enter a valid printer number.")

if __name__ == '__main__':
    main()
