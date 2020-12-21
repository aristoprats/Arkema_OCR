import tesserocr
from pdf2image import convert_from_path
from PIL import Image
import os
import shutil

###### GLOBAL VARS ######
GLOBAL_poppler_path = r'dependencies\\poppler_home\\bin'
GLOBAL_temp_path = r'temp'
GLOBAL_archive_path = r'Archive'
GLOBAL_production_path = r'To_Scan'
GLOBAL_desired_headers = ['Functional Location', 'Equipment', 'Main Work Center', 'Oper. Short Text','Section','Partner Number']
#########################



def run_precheck():
    directory_list = os.listdir()

    if GLOBAL_temp_path in directory_list:
        shutil.rmtree(GLOBAL_temp_path)
    os.mkdir(GLOBAL_temp_path)

    if GLOBAL_archive_path not in directory_list:
        os.mkdir(GLOBAL_archive_path)
    
    if GLOBAL_production_path not in directory_list:
        os.mkdir(GLOBAL_production_path)
    

def run_cleanup():
    shutil.rmtree(GLOBAL_temp_path)

def gather_scannables():
    scannables = []
    for scannable in os.listdir(GLOBAL_production_path + '\\'):
        scannables += [scannable]
    
    return scannables

def create_jpgs(to_convert, scan_index=0, dpi=300):
    converted = convert_from_path(pdf_path= GLOBAL_production_path + '\\' + to_convert, dpi=dpi, poppler_path=GLOBAL_poppler_path)
    temp_filename = GLOBAL_temp_path + '\\' + 'temp_image_%d_.jpg' % scan_index
    converted[0].save(temp_filename, 'JPEG')
    return temp_filename

def run_ocr(ocr_img):
    pil_image = Image.open(ocr_img)
    return tesserocr.image_to_text(pil_image)
    
def rejoin_lines(char_array):
    string_array = []
    
    char_idx = 0
    string_line = ''
    while char_idx < len(char_array):
        if char_array[char_idx] == '\n':
            string_array += [string_line]
            string_line = ''
        else:
            string_line += char_array[char_idx]
        char_idx += 1

    return string_array
        
def search_for_PM(intake):
    found = 'FALSE'

    for line in intake:
        if 'EG-MAIN' in line:
            pm_str = ''
            idx = line.find('EG-MAIN')
            for char in line[idx + len('EG-MAIN-001-'):]:
                if char != ' ':
                    pm_str += char
                else:
                    found = pm_str
                    break
    
    return [found]

def parse_text(intake, archive_ID=-1):
    joined_text = rejoin_lines(intake)
    line_to_write = []

    for header in GLOBAL_desired_headers:
        located = False
        for line in joined_text:
            if header in line:
                try:
                    line_to_write.append(line.split(':')[1])
                    located = True
                except:
                    pass
        if not located:
            line_to_write.append('-')

    line_to_write = [archive_ID] + search_for_PM(joined_text) + line_to_write

    return line_to_write



def main():

    run_precheck()

    production_list = gather_scannables()
    
    scan_idx = 0
    for scannable in production_list:
        scan_idx += 1
        ocr_img = create_jpgs(scannable,scan_index=scan_idx)
        raw_text = run_ocr(ocr_img)
        parsed = parse_text(raw_text)
        print(parsed)

    


    
        
    
    
    run_cleanup()

if __name__ == '__main__':
    main()