import os
import pdfplumber
import openpyxl


# Function to get file names
def get_file_names():
    folder_path = os.getenv("MY_FILE_PATH")
    
    # List all files in the folder
    file_names = os.listdir(folder_path)

    # Filter out directories and only keep files
    file_names = [f for f in file_names if os.path.isfile(os.path.join(folder_path, f))]

    files = []
    print(file_names)
    for file in file_names:
        file_path = os.path.join(folder_path, file)
        files.append(file_path)
    return files

# Function to read PDF and extract text
def extract_pdf_text(file_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

# Function to parse specific fields from extracted PDF text
def parse_data(text):
    # Replace with the specific fields you want to extract
    # Example: Assuming you're extracting "Name" and "Date" fields from the text
    data = []
    lines = text.split('\n')
    for line in lines:
        if "Order" in line:
            po_number = line.split()
            data.append(po_number[-1])
        elif "Project name:" in line:
            project_name = line.split()
            data.append(project_name[-1].split(":")[-1])
        elif "Ordered:" in line:
            order_date = line.split(":")
            data.append(order_date)
        elif "Total price:" in line:
            total_price = line.split()
            data.append(total_price[-2])
    return data

# Function to write extracted data into an Excel file
def write_to_excel(data, output_file_path, k):

    file_path = output_file_path

    try:
        # If the file exists, load the workbook
        workbook = openpyxl.load_workbook(file_path)
    
    except FileNotFoundError:
        # If the file does not exist, create a new workbook
        workbook = openpyxl.Workbook()
    
    sheet = workbook.active

    # Write data rows
    sheet[f"A{k}"] = f"Purchase Order No.{data[0]} dated {data[2]}, {data[1]}"
    sheet[f"E{k}"] = float(data [3])
    sheet[f"C{k}"] = float(data [3])
        

    workbook.save(output_file_path)
    print(f"Data written to {output_file_path}")

def iterate_over_files():
    pass


# Main function to handle the entire process
def main():
    
    output_file_path = os.getenv("output_file_path")          # Path to the output Excel file

    
    
    # Write parsed data to an Excel file
    array_of_files = get_file_names()
    k = 22    
    for pdf_file_path in array_of_files:
        # Extract text from the PDF
        try:
            pdf_text = extract_pdf_text(pdf_file_path)
            data = parse_data(pdf_text)
            data_refact = data[-2].split(":")[-1]
            data[-2] = data_refact

            print(data)
            write_to_excel(data, output_file_path, k)
            k+=1
        except Exception as e:
             print(f"Error: {e} - The file could not be parsed.")

        # Parse specific fields from the text
        

       

# Run the main function
if __name__ == "__main__":
    main()
