import pytesseract as pt
import pdf2image
import os
import shutil

def run_ocr(filename, dpi=500):
    pages = pdf2image.convert_from_path(pdf_path=filename, dpi=dpi, thread_count=(os.cpu_count()*1.5))

    raw_string_split = pt.image_to_string(pages[0]).split('\n')
    return [x for x in raw_string_split if x != '']
    
def setup_excel_row(ocr_string, desired_headers):
    line_to_write = []
    for header in desired_headers:
        located = False
        for line in ocr_string:
            if header in line:
                try:
                #print(header, '|',line.split(':')[1])
                    line_to_write.append(line.split(':')[1])
                    located = True
                except:
                    pass
        if not located:
            #print(header, '|', '----')
            line_to_write.append('-')
    
    return line_to_write
                
def write_to_excel(archive, list_to_write):
    write_string = ''
    for item in list_to_write:
        write_string += item + ','
    
    archive.write(write_string + '\n')

def get_last_ID(last_archive):
    with open(last_archive, 'rb') as f:
        f.seek(-2,2)
        while f.read(1) != b"\n":
            f.seek(-2,1)
        return f.read().decode('utf-8').split(',')[0]

def search_for_PM(ocr_string):
    PM_str = ['FALSE']

    for item in ocr_string:
        if 'EG-MAIN' in item:
            pm_str = ''
            idx = item.find('EG-MAIN')
            for char in item[idx:]:
                if char != ' ':
                    pm_str += char
                else:
                    PM_str = [pm_str]
                    break
    return PM_str

def perform_ocr_work(archive, revnum, desired_headers, last_archive):
    files_to_ocr = add_files_to_parse()
    #print(files_to_ocr)
    
    if last_archive == -1:
        write_to_excel(archive, ['Archive ID', 'EG-MAIN?'] + desired_headers)
        ID_assignment = '1'
    else:
        ID_assignment = str(int(get_last_ID(last_archive)) + 1)

    for file in files_to_ocr:
        ocr_string = run_ocr(file)
        #print(*ocr_string, sep='| \n')
        #print("\n" * 5)
        row_to_excel = [ID_assignment] + search_for_PM(ocr_string) + setup_excel_row(ocr_string, desired_headers)
        #print(*row_to_excel, sep='\n')
        write_to_excel(archive, row_to_excel)
        shutil.move(file, f'ARCHIVE/{ID_assignment}.pdf')
        ID_assignment = str(int(ID_assignment) + 1)
        

def add_files_to_parse():
    files_to_ocr = []
    for file in os.listdir('.'):
        if '.pdf' in file:
            files_to_ocr.append(file)
    return files_to_ocr

def main():
    lead_name = 'PM_PAPERWORK_ARCHIVE_'
    last_revnum = -1
    for file in os.listdir():
        if lead_name in file and ".csv" in file:
            if int(file[len(lead_name):-4]) > last_revnum:
                last_revnum = int(file[len(lead_name):-4])
    
    archive_name = (lead_name + str(last_revnum + 1) + '.csv')
    if last_revnum >= 0:
        last_archive_name = (lead_name + str(last_revnum) + '.csv')
        shutil.copyfile(last_archive_name, archive_name)
    else:
        last_archive_name = -1


    desired_headers = ['Functional Location', 'Equipment', 'Main Work Center', 'Oper. Short Text','Section','Partner Number']
    archive = open(archive_name, 'a')
    perform_ocr_work(archive, last_revnum + 1, desired_headers, last_archive_name)
    archive.close()

if __name__ == '__main__':
    main()