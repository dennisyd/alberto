import os
import win32com.client
import pythoncom
import time
win32com.client.gencache.EnsureDispatch('Excel.Application')


def replace_formula_and_fill_down(file_path):
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        try:
            print(f"Opening workbook: {file_path}")
            workbook = excel.Workbooks.Open(os.path.abspath(file_path))
            time.sleep(1)  # Give Excel time to fully open the workbook

            # Iterate through all sheets that have '2024' in their names
            sheets_processed = 0
            for sheet in workbook.Worksheets:
                if '2024' in sheet.Name:
                    print(f"Processing sheet: {sheet.Name}")

                    # Define the existing formula
                    formula = '''=IFERROR(
IF(AND(L9="ARGAWALA", I9="PENDING", NOT(OR(ISNUMBER(SEARCH("TCD", F9)), ISNUMBER(SEARCH("ALLERGY", F9)), ISNUMBER(SEARCH("AFT", F9))))), -4,
IF(AND(L9="ARGAWALA", I9="PENDING"), 0.7*I9,
IF(AND(L9="ARGAWALA", OR(ISNUMBER(SEARCH("TCD", F9)), ISNUMBER(SEARCH("ALLERGY", F9)), ISNUMBER(SEARCH("AFT", F9)))), 0.7*I9,
IF(L9="ARGAWALA", 0.7*I9 -4,
IF(AND(I9="PENDING", OR(L9="AUSUBEL", L9="MOREHOUSE", L9="MITCHEL", L9="FEFER", L9="RIGNEY", L9="ZAMBITO", L9="KEIL", L9="BONHEIM",L9="LISANN", L9="RAMACHANDRAN", L9="SAHAI", L9="TRAZZERA", L9="FINKELSTEIN")), -40,
IF(OR(L9="SAHAI", L9="FINKELSTEIN", L9="LISANN", L9="TRAZZERA", L9="BONHEIM", L9="RAMACHANDRAN", L9="KEIL",L9="MITCHEL"), 0.75*I9,
IF(OR(L9="MODI", OR(L9="WALLERSON", L9="WALLERSON/DAVE")), 0.7*I9,
""))))))),"")'''

                    try:
                        # Set the formula in cell K9
                        sheet.Range("K9").Formula = formula
                        print(f"Formula set in K9 of sheet: {sheet.Name}")

                        # Fill the formula down to K3020
                        sheet.Range("K9:K3020").FillDown()
                        print(f"Formula filled down to K3020 in sheet: {sheet.Name}")
                        sheets_processed += 1

                    except Exception as cell_error:
                        print(f"Error with cell operations in sheet {sheet.Name}: {str(cell_error)}")
                        raise

            if sheets_processed > 0:
                print(f"Total sheets processed: {sheets_processed}")
            else:
                print("No sheets found with '2024' in the name.")

            # Save and overwrite the original file
            print(f"Saving and overwriting: {file_path}")
            workbook.Save()
            workbook.Close()
            print("Process completed successfully")
            return True

        finally:
            excel.Quit()

    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()

def main():
    # Directory path
    directory = r"C:\\Users\\denni\\Downloads\\Alberto 2024\\2025"

    # Process all .xlsx files in the directory
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            input_file = os.path.join(directory, filename)

            print(f"\nProcessing: {filename}")

            if replace_formula_and_fill_down(input_file):
                print("Successfully replaced formula and filled down")
            else:
                print("Failed to replace formula and fill down")

if __name__ == "__main__":
    main()