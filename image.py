import time

from PIL import Image
import pytesseract
from pathlib import Path


# pytesseract.pytesseract.tesseract_cmd = r'C:\Users\a.soyunjaliyev\AppData\Local\Programs\Tesseract-OCR'
image_path=Path(r'D:\BYTK_Facturation\7. MT\image')

start = time.perf_counter()
for i, image in enumerate(image_path.glob('*.jpg'), start=1):
    print(pytesseract.image_to_string(Image.open(image)))
    break

finish = time.perf_counter()
print('Time spent {} seconds.'.format(round(finish - start, 0)))