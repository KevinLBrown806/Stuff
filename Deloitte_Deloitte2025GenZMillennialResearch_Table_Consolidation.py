import os
import time
import win32com.client

def convert_to_xlsm(file_path, excel):
    abs_path = os.path.abspath(file_path)
    # Open the workbook
    workbook = excel.Workbooks.Open(abs_path)
    
    # Create new file path with .xlsm extension (in the same folder as the input file)
    new_file = os.path.splitext(abs_path)[0] + ".xlsm"
    
    # Disable alerts so that Excel overwrites without prompting
    excel.DisplayAlerts = False
    workbook.SaveAs(new_file, FileFormat=52)
    print(f"Converted '{file_path}' to '{new_file}'")
    
    return new_file, workbook

def embed_bas_files(workbook, bas_folder):
    # List all .bas files in the specified folder
    bas_files = [os.path.join(bas_folder, f)
                 for f in os.listdir(bas_folder)
                 if f.lower().endswith('.bas')]
    
    if not bas_files:
        print(f"No .bas files found in folder: {bas_folder}")
        return
    
    # Import each .bas file into the workbook's VBA project
    for bas_file in bas_files:
        try:
            workbook.VBProject.VBComponents.Import(bas_file)
            print(f"Embedded '{bas_file}' into '{workbook.Name}'")
        except Exception as e:
            print(f"Failed to embed '{bas_file}' into '{workbook.Name}': {e}")

def get_macro_name(bas_path):
    # Extract macro name from the file name (without extension)
    return os.path.splitext(os.path.basename(bas_path))[0]

def run_macros(workbook, excel, macro_paths):
    # Run macros in the specified order with a pause between them
    for macro in macro_paths:
        macro_name = get_macro_name(macro)
        try:
            # Activate the workbook to ensure it is the active one
            workbook.Activate()
            # Run the macro using the workbook's name (enclosed in single quotes in case of spaces)
            excel.Application.Run(f"'{workbook.Name}'!{macro_name}")
            print(f"Ran macro: {macro_name}")
            # Pause to allow Excel to finish processing before running the next macro
            time.sleep(3)
        except Exception as e:
            print(f"Failed to run macro '{macro_name}': {e}")

def convert_to_xlsx(workbook, input_file, output_folder, excel):
    # Build the output file path: same file name as the input, but with .xlsx extension in the output folder.
    base_name = os.path.basename(input_file)
    name_without_ext = os.path.splitext(base_name)[0]
    new_output_file = os.path.join(output_folder, name_without_ext + ".xlsx")
    try:
        # FileFormat 51 saves as XLSX (which does not support macros)
        workbook.SaveAs(new_output_file, FileFormat=51)
        print(f"Converted '{input_file}' to XLSX '{new_output_file}'")
    except Exception as e:
        print(f"Failed to convert '{input_file}' to XLSX: {e}")
    return new_output_file

def main(input_folder, bas_folder, output_folder):
    # Filter to include only XLSX files that do not start with "~$" (temporary files)
    xlsx_files = [os.path.join(input_folder, f)
                  for f in os.listdir(input_folder)
                  if f.lower().endswith('.xlsx') and not f.startswith("~$")]
    
    if not xlsx_files:
        print(f"No XLSX files found in folder: {input_folder}")
        return

    # Ordered list of macro BAS file paths (the order in which to run them)
    macro_order = [
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\ConsolidateSheets01.bas",
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\CreateTableOfContents02.bas",
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\CopyDataFromTOCToTableOfContentsWithHyperlinksPreserved03.bas",
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\DeleteColumnB04.bas",
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\DeleteAllSheetsExceptSpecific05.bas",
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\CopyDataFromAnotherWorkbook06.bas",
        r"C:\Users\kbrown\Desktop\Deloitte_2024\VBA\CleanTOC07.bas"
    ]
    
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Run Excel in the background
    excel.DisplayAlerts = False

    for file in xlsx_files:
        try:
            # Convert to XLSM and open the workbook
            new_file, workbook = convert_to_xlsm(file, excel)
            # Embed BAS files into the workbook's VBA project
            embed_bas_files(workbook, bas_folder)
            # Run the macros in the specified order
            run_macros(workbook, excel, macro_order)
            # Convert back to XLSX into the output folder
            convert_to_xlsx(workbook, file, output_folder, excel)
            workbook.Close(SaveChanges=True)
        except Exception as e:
            print(f"Error processing '{file}': {e}")
    
    excel.DisplayAlerts = False
    excel.Quit()

if __name__ == "__main__":
    input_folder = input("Enter the folder path containing XLSX files: ")
    bas_folder = input("Enter the folder path containing .bas files: ")
    output_folder = input("Enter the output folder path for XLSX files: ")
    main(input_folder, bas_folder, output_folder)
