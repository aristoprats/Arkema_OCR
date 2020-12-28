##### Imports #####
#   Most of these imports only use 1-4 functions, but importing full libraries because minimal performance loss and it makes reading easier
import tesserocr
from pdf2image
from PIL import Image
import os              
import shutil          
import pandas as pd
import openpyxl         # Don't need, including to have engine available post compilation
from time import sleep        # To Do: Remove after testing
from datetime import datetime # To Do: Remove after testing
#########################

###### GLOBAL VARS ######
# Globals - paths to file system. Using globals to ensure consistency and only instantiate variables once
## KEY NOTE: all paths except poppler are 1 directory above program
GLOBAL_poppler_path = r'dependencies\\poppler_home\\bin'
GLOBAL_temp_path = r'..\\temp'
GLOBAL_archive_path = r'..\\Archive'
GLOBAL_production_path = r'..\\To_Scan'
GLOBAL_excel_sheet_prefix = r'..\\PM_Index_'

# Globals - information to pull from PMs. Using global to ensure consistency and only instantiate variables once
GLOBAL_desired_headers = ['ID#', 'EG-MAIN-001', 'Functional Location', 'Equipment', 'Main Work Center', 'Oper. Short Text','Section','Partner Number']
#########################

class Excel_to_write_df:
    ''' A class to represent the excel data from the last time the program was run. 
    This class is not critical but was created to bundle methods relating to excel IO and updates

    Attributes
    -------
    new_name: str 
        filepath and name for next revision of excel index
    _df     : pandas dataframe
        hidden attribute containing the loaded dataframe 
    _locked : bool
        hidden attribute to control shared data modification during multiprocessing
    
    Methods
    ------
    get_last_id_num() -> returns int
        returns the archive ID# of last entry in loaded excel dataframe
    get_lock_stats()  -> returns bool
        returns whether or not _df is currently being modified
    flip_lock()       -> returns bool
        flips lock status and returns current lock status
    update_primary_df -> returns none
        concats new row to index excel dataframe
    write_next_excel  -> returns none
        writes current index excel dataframe to disk
    display_df        -> return dataframe object
        accesor function for hidden attribute _df
    '''

    def __init__(self, last_excel):
        '''Constructor function for excel dataframe object. Takes only version number of last index excel'''
        self.new_name = GLOBAL_excel_sheet_prefix + str(last_excel + 1) + '.xlsx'
        self._df = pd.read_excel(GLOBAL_excel_sheet_prefix + str(last_excel) + '.xlsx', engine='openpyxl')
        self._locked = False

    def get_last_idnum(self):
        '''Returns last archived ID# in index excel'''
        return int(self._df.tail(1)['ID#'])
    
    def get_lock_stats(self):
        '''Returns status of dataframe. True if being modified, False if avaibale to modify'''
        return self._locked
    
    def flip_lock(self):
        '''Inverts status of dataframe lock'''
        self._locked = not self._locked
        return self._locked

    def update_primary_df(self, new_df_row):
        '''Takes new OCR'd data as a dataframe row and concatenates to existing row'''
        # Poorly done way of managing shared resources :( 
            # TO DO: revisit if implementing multiprocessing
        while self._locked:
            sleep(2)
        
        self.flip_lock()
        temp = pd.concat([self._df, new_df_row], ignore_index=True)
        self._df = temp
        #   Needs to be written as two lines, updating inline was giving issues?
        self.flip_lock()
        
    def write_next_excel(self):
        '''Write current index excel dataframe to disk, does not write out df index'''
        self._df.to_excel(self.new_name, index=False)

    def display_df(self):
        '''Accessor function to visualize index excel dataframe. For debugging'''
        temp = self._df
        return temp

def run_precheck():
    ''' Single use pre-OCR prep and staging function. Scans output directory and generates/removes file structure as needed
        
        Returns:
            int - most recent revision to index excels
    '''
    directory_list = os.listdir(r'..\\')
    
    # clean any pre-existing temp directory
    if GLOBAL_temp_path[4:] in directory_list:
        shutil.rmtree(GLOBAL_temp_path)
    os.mkdir(GLOBAL_temp_path)

    # Create an archive directory if not present
    if GLOBAL_archive_path[4:] not in directory_list:
        os.mkdir(GLOBAL_archive_path)
    
    # Create a production path if doesn't exist
    #    also exits as this means no files are loaded
    if GLOBAL_production_path[4:] not in directory_list:
        os.mkdir(GLOBAL_production_path)
        exit
    
    # Scan for latest index version
    list_existing_indexes = []
    len_to_ignore = len(GLOBAL_excel_sheet_prefix[4:])
    for file in directory_list:
        if '.xlsx' in file:
            list_existing_indexes.append(int(file[len_to_ignore:-5]))
    
    return max(list_existing_indexes) #return most recent version#
    
def run_cleanup():
    ''' Removes temp directory and all files that were just added

        Returns:
            none
    '''
    shutil.rmtree(GLOBAL_temp_path)
    '''
    for file in os.listdir(r'To_Scan'):
        os.remove(file)

    ''' #To Do: Remove this once testing finalized

def gather_scannables():
    ''' Scan production directory and collect filepaths to all pdfs

        Returns:
            Scannables (list): List of filepaths relative to program location
    '''
    scannables = []
    for scannable in os.listdir(GLOBAL_production_path + '\\'):
        scannables += [scannable]
    
    return scannables

def create_jpgs(to_convert, scan_index=0, dpi=300):
    ''' Converts existing PDFs to JPEGs. Needed for tesseract-ocr

        Parameters:
            to_convert (str) - filepath of pdf, relative to program location
            scan_index (int) - archive number to assign this pdf, defaults to 0
            dpi        (int) - resolution to convert image into jpeg, defaults to 300
    '''
    converted = pdf2image.convert_from_path(pdf_path=to_convert, dpi=dpi, poppler_path=GLOBAL_poppler_path, thread_count=3)
    temp_filename = GLOBAL_temp_path + '\\' + 'temp_image_%d_.jpg' % scan_index
    converted[0].save(temp_filename, 'JPEG')
    return temp_filename

def run_ocr(ocr_img):
    ''' Convert JPEG to PIL Image and then run tesseract-ocr
        Parameters:
            ocr_img (str) - filepath to jpeg file
        Returns: 
            (character array) - returns OCR reading as character array
    '''
    pil_image = Image.open(ocr_img)
    return tesserocr.image_to_text(pil_image)
    
def rejoin_lines(char_array):
    ''' Takes a character array and stitches lines together by searching for line returns
        Parameters:
            char_array (char_array) - char_array from tesseract
        Returns:
            string_array (list (str)) - str_array with each line as a string
    '''
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
    ''' Searches string array to check if 'EG-MAIN' procedure.
        Parameters:
            intake (list (str)) - str_array from rejoin_lines()
        Returns:
            found  ( [str] )    - single item list with either the PM procedure letter or 'FALSE'
    '''
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
    ''' Parse pdf text by checking line by line for headers in GLOBAL_desired_headers. Easily configurable
        Parameters:
            intake (list (str)) - str_array from rejoin_lines()
            archive_ID    (int) - index excel ID# for this pdf
        Return:
            line_to_write (list (str)) - data for all headers with matching index to GLOBAL_desired_heards
    '''
    joined_text = rejoin_lines(intake)
    line_to_write = []

    # Yes this is a slow nested for loop, but there are only a few relevent headers and only OCR'ing first page of sparse PDFs
    ## Can Improve this, but performance is already well above what is needed so leaving for now
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
    ''' Convert a data list to a row dataframe
        Parameters:
            parsed_list (list (str)) - str_array from parse_text
        Return:
            (pandas dataframe) - row dataframe with all scraped data
    '''
        return pd.DataFrame([parsed_list], columns=GLOBAL_desired_headers)



def main():

    # Pre OCR checklist
    last_ver = run_precheck()
    next_df = Excel_to_write_df(last_ver)
    production_list = gather_scannables()
    scan_idx = next_df.get_last_idnum() + 1

    # Main program loop. Set up to be easily multiprocessable, however single threaded performance was beyond good enough so mp work was put off for now
    for scannable in production_list:
        # loop iteration - prep phase
        To_Scan = GLOBAL_production_path + '\\' + scannable
        new_destination = GLOBAL_archive_path + '\\' + str(scan_idx) + '.pdf'
        shutil.copyfile(To_Scan, new_destination)
        # loop iteration - OCR phase
        ocr_img = create_jpgs(To_Scan,scan_index=scan_idx)
        raw_text = run_ocr(ocr_img)
        # loop iteration - scrape data from OCR data
        parsed = parse_text(raw_text, scan_idx)
        parsed_frame = list_to_dataframe(parsed)
        # loop iteration - update master dataframe and step forward
        next_df.update_primary_df(parsed_frame)
        scan_idx += 1
    
    # Export updated index excel df by writing out
    next_df.write_next_excel()      
    
    # Last function to clean up any remaining temp files
    run_cleanup()

if __name__ == '__main__':
    startTime = datetime.now()

    # To Do: Remove in final version, just for debugging by creating an artifically large dataset
    number_of_trials = 1
    for _ in range(0,number_of_trials):
        main()
    
    end = datetime.now()

    print("##### Runtime Statistics #####")
    print(f"Total time for {5*number_of_trials} samples : \t{(end - startTime)/60} minutes")
    print(f"Average time per sample: \t{(end - startTime)/60/(5*number_of_trials)} minutes per sample")