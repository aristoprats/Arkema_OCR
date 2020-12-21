#from tesserocr import PyTessBaseAPI
import tesserocr
from pdf2image import convert_from_path
from PIL import Image

GLOBAL_poppler_path = r'poppler_home\\bin'
GLOBAL_temp_path = r'temp\\'

sample_file = 'sample1.pdf'
sample_img = convert_from_path(pdf_path=sample_file, dpi=300, poppler_path=GLOBAL_poppler_path, output_folder='temp', fmt='jpg')
sample_img[0].save('TEST.jpg', 'JPEG')


'''
with PyTessBaseAPI() as api:
    api.SetImageFile('TEST.jpg')
    print(api.GetUTF8Text())
'''

pil_image = Image.open('TEST.jpg')

print(tesserocr.image_to_text(pil_image))