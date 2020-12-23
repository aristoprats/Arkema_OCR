import tesserocr
from pdf2image import convert_from_path
from PIL import Image
import os
import shutil
import pandas as pd
import openpyxl
from time import sleep

#hidden dependencoes
# openpyxl - pandas reading excel 

###### GLOBAL VARS ######
GLOBAL_poppler_path = r'dependencies\\poppler_home\\bin'
GLOBAL_temp_path = r'temp'
GLOBAL_archive_path = r'Archive'
GLOBAL_production_path = r'To_Scan'
GLOBAL_desired_headers = ['ID#', 'EG-MAIN-001', 'Functional Location', 'Equipment', 'Main Work Center', 'Oper. Short Text','Section','Partner Number']
GLOBAL_excel_sheet_prefix = 'PM_Index_'
#########################

class Excel_to_write_df:
    def __init__(self, last_excel):
        self.new_name = GLOBAL_excel_sheet_prefix + str(last_excel + 1) + '.xlsx'
        self._df = pd.read_excel(GLOBAL_excel_sheet_prefix + str(last_excel) + '.xlsx', engine='openpyxl')
        self._locked = False

    def get_last_idnum(self):
        return int(self._df.tail(1)['ID#'])
    
    def get_lock_stats(self):
        return self._lock
    
    def flip_lock(self):
        self._locked = not self._locked
        return self._locked

    def update_primary_df(self, new_df_row):

        # Poorly done way of managing shared resources :( 
            # TO DO: revisit when understand multiprocessing and shared data management better
        while self._locked:
            sleep(2)
        
        self.flip_lock()
        temp = pd.concat([self._df, new_df_row], ignore_index=True)
        self._df = temp
        self.flip_lock()
        

    def write_next_excel(self):
        self._df.to_excel(self.new_name, index=False)

    def display_df(self):
        temp = self._df
        return temp


    

def run_precheck():
    directory_list = os.listdir()

    # clean any pre-existing temp directory
    if GLOBAL_temp_path in directory_list:
        shutil.rmtree(GLOBAL_temp_path)
    os.mkdir(GLOBAL_temp_path)

    # Create an archive directory if not present
    if GLOBAL_archive_path not in directory_list:
        os.mkdir(GLOBAL_archive_path)
    
    # Create a production path if doesn't exist
    #    also exits as this means no files are loaded
    if GLOBAL_production_path not in directory_list:
        os.mkdir(GLOBAL_production_path)
        exit
    
    # Scan for latest index version
    list_existing_indexes = []
    len_to_ignore = len(GLOBAL_excel_sheet_prefix)
    for file in os.listdir():
        if '.xlsx' in file:
            list_existing_indexes.append(int(file[len_to_ignore:-5]))
    
    return max(list_existing_indexes) #return most recent version#
    
def run_cleanup():
    shutil.rmtree(GLOBAL_temp_path)
    '''
    for file in os.listdir(r'To_Scan'):
        os.remove(file)

    ''' #To Do: Remove this once testing finalized

def gather_scannables():
    scannables = []
    for scannable in os.listdir(GLOBAL_production_path + '\\'):
        scannables += [scannable]
    
    return scannables

def create_jpgs(to_convert, scan_index=0, dpi=300):
    converted = convert_from_path(pdf_path=to_convert, dpi=dpi, poppler_path=GLOBAL_poppler_path, thread_count=3)
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
    found = False

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

    for header in GLOBAL_desired_headers[2:]:
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

    # Implementation as a list
    line_to_write = [archive_ID] + search_for_PM(joined_text) + line_to_write

    return line_to_write

def list_to_dataframe(parsed_list):
    # return pd.DataFrame(np.array(parsed_list).reshape(-1,len(parsed_list)), columns=GLOBAL_desired_headers)

    return pd.DataFrame([parsed_list], columns=GLOBAL_desired_headers)



def main():

    last_ver = run_precheck()
    next_df = Excel_to_write_df(last_ver)

    production_list = gather_scannables()



    scan_idx = next_df.get_last_idnum()
    for scannable in production_list:
        To_Scan = GLOBAL_production_path + '\\' + scannable
        new_destination = GLOBAL_archive_path + '\\' + str(scan_idx) + '.pdf'
        shutil.copyfile(To_Scan, new_destination)
        scan_idx += 1
        ocr_img = create_jpgs(To_Scan,scan_index=scan_idx)
        raw_text = run_ocr(ocr_img)
        parsed = parse_text(raw_text, scan_idx)
        parsed_frame = list_to_dataframe(parsed)
        next_df.update_primary_df(parsed_frame)
    

    
    next_df.write_next_excel()      
    
    run_cleanup()

if __name__ == '__main__':
    main()