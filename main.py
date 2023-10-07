import subprocess
import configparser
import os

def run_process():
    # Load configuration
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Run email processor
    vbs_script_path = config['Paths']['VBScript']
    excel_file = config['Paths']['ExcelFile']
    macro_name = config['Other']['ExcelMacroName']
    subprocess.run(["cscript.exe", vbs_script_path, excel_file, macro_name])

    # Run PDF extractor
    pdf_extractor_path = config['Paths']['PDFExtractor']
    processed_folder = config['Paths']['ProcessedFolder']
    subprocess.run([pdf_extractor_path, processed_folder])  # Updated line

if __name__ == "__main__":
    run_process()
