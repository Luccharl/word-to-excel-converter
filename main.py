import os, docx, openpyxl

def word_to_excel(folder_path, file_name):
    
    # New excel workbook
    wb = openpyxl.Workbook()

    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            
            # Path to the .docx file
            doc_path = os.path.join(folder_path, filename)
            
            # Open the word document
            doc = docx.Document(doc_path)
            
            for i, table in enumerate(doc.tables):
                # New sheet for the table
                sheet_name = f"{filename}_{i+1}"
                ws = wb.create_sheet(sheet_name)
                
                for j, row in enumerate(table.rows):
                    for k, cell in enumerate(row.cells):
                        # Add data to sheet
                        ws.cell(row=j+1, column=k+1).value = cell.text
    
    # Save the workbook
    wb.save(f"{file_name}.xlsx")
    
    return f"{file_name}.xlsx successfully created on {folder_path}."

if __name__ == "__main__":
    
    # Folder containing .docx files
    folder_path = input('Type the path of the folder: ')

    # Name of the excel workbook to be created
    file_name = input('Type the filename of the excel file to be created: ')

    print(word_to_excel(folder_path, file_name))