import os  #used to interact with files in a directory
import re   #used for extracting information using regular expressions
from datetime import datetime # for date formatting
from striprtf.striprtf import rtf_to_text # for converting RTF to plain text
import win32com.client  # for interacting with Windows COM objects



# Input and output folders
documents = r'your file path' # Folder containing unorganized RTF files
Organized_Documents = r'Final file path'  # Folder where organized files will be stored
missed_files = r'hold over file path' # Folder for files that could not be processed



#initialize the word application
word = win32com.client.Dispatch("Word.Application") # Creates a Word application object

# function that formats dates
def format_date(date_str):
    try:
        return datetime.strptime(date_str, '%m/%d/%Y').strftime('%Y-%m-%d') # Convert date from MM/DD/YYYY to YYYY-MM-DD format
    except ValueError: 
        return None

# Main organization function
def organize_rtf_files():
    
    for filename in os.listdir(documents):   # Loop through each file in the documents folder
        if filename.lower().endswith('.rtf'): # Check if the file is an RTF file and ignores any that are not
            file_path = os.path.join(documents, filename) # Construct the full file path by joining folder path and filename to get full path
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file: # reads the file as a plain text utf-8 and ignores errors
                raw_rtf = file.read() # reads the raw RTF content
                content = rtf_to_text(raw_rtf)

            # Extract info from file content using regular expressions
            name_match = re.search(r'Patient Name:\s*(\w+)\s+(\w+)', content) # Extracts first and last name
            dob_match = re.search(r'DOB:\s*(\d{2}/\d{2}/\d{4})', content) # Extracts date of birth
            doc_type_match = re.search(r'(\Bmaster.+im)', file_path) # Extracts document type
            if not doc_type_match:  # If the document type is not found, try a different pattern
                doc_type_match = re.search(r'(\Bchart.+note)', file_path) # Extracts document type from content
            enc_date_match = re.search(r'Date:\s*(\d{2}/\d{2}/\d{4})', content) # Extracts encounter date

            if name_match and dob_match and doc_type_match and enc_date_match:
                first_name = name_match.group(1).capitalize() # Capitalizes the first name
                last_name = name_match.group(2).capitalize() # Capitalizes the last name
                dob = format_date(dob_match.group(1)) # Formats the date of birth
                enc_date = format_date(enc_date_match.group(1)) # Formats the encounter date
                doc_type = doc_type_match.group(0).strip().replace(' ', '_')
                

                
                


                if not all([first_name, last_name, dob, doc_type , enc_date]): # Checks if any of the extracted fields are missing due to not having the specified content i.e. dob last_name etc
                    print(f"Skipping due to missing fields")

                    continue 

                # Folder structure: Organized/S/Smith, John_1980-05-12/
                first_letter = last_name[0].upper()
                patient_folder = os.path.join(Organized_Documents, first_letter, f"{last_name}, {first_name}_{dob}") # Construct the patient folder path 
                os.makedirs(patient_folder, exist_ok=True) # Create the patient folder if it doesn't exist

                
                
                


                try:
                    # Define the new PDF filename
                    new_pdf_filename = f"{doc_type}_{enc_date}.pdf"
                    new_pdf_file_path = os.path.join(patient_folder, new_pdf_filename) # Construct the new file path with the new filename
                    missed_file_path = os.path.join(missed_files, filename) # Path for missed files


                   
                    

                    doc = word.Documents.Open(file_path) # Open the RTF file in Word
                    doc.SaveAs2(new_pdf_file_path, FileFormat=17) # Save the document as a PDF (FileFormat=17 is for PDF)
                    doc.Close() # Close the document before copying
                    
                    
                    print(f"PDF created")
                except Exception as e:
                    print(f"Error converting to PDF: {e}")
                    word.Quit() # Ensure Word application is closed if there is an error 
               
                
                print(f"Moved ")
            else:
                print(f"Missing information. relocated file")
                #os.makedirs(missed_files, exist_ok=True) # Create the missed files folder if it doesn't exist
                os.rename(file_path, missed_file_path) # Move the file to the missed files folder


                
    word.Quit() # Ensure Word application is closed after processing all files
    print("Closed Word application")         
organize_rtf_files()


#TODO 1.: Add logging to a file instead of printing to console

