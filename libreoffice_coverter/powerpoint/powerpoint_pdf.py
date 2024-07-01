import os
import subprocess

def convert_powerpoint_to_pdf(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    
    for filename in os.listdir(input_folder):
        if filename.endswith('.pptx') or filename.endswith('.ppt'):
            input_file = os.path.join(input_folder, filename)
            output_file = os.path.join(output_folder, os.path.splitext(filename)[0] + '.pdf')
            
            # Specify the full output file path where LibreOffice will save the PDF
            output_pdf_file = os.path.join(os.getcwd(), os.path.splitext(filename)[0] + '.pdf')
            
            result = subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.getcwd(), input_file], capture_output=True, text=True)
            if result.returncode == 0:
                # Move the converted PDF file to the output folder
                if os.path.exists(output_pdf_file):
                    os.rename(output_pdf_file, output_file)
                    print(f"Successfully converted {filename} to PDF.")
                else:
                    print(f"Converted {filename}, but PDF file not found in {os.getcwd()}.")
            else:
                print(f"Failed to convert {filename}.")
                print(result.stderr)

# Example usage
input_folder = './input_folder_powerpoint'
output_folder = './output_folder_powerpoint'

convert_powerpoint_to_pdf(input_folder, output_folder)