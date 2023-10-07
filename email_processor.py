import subprocess
import configparser

def run_vba_macro():
    # Load configuration
    config = configparser.ConfigParser()
    config.read('config.ini')

    excel_file = config['Paths']['ExcelFile']
    vbs_script_path = config['Paths']['VBScript']
    macro_name = config['Other']['ExcelMacroName']

    # Run VBScript with arguments
    subprocess.run(["cscript.exe", vbs_script_path, excel_file, macro_name])

# Uncomment the line below to run the function
run_vba_macro()

