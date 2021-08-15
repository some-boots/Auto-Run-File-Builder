import pdfminer
from pdfminer import high_level

run_file_laparams = pdfminer.layout.LAParams(char_margin=1000)

text = pdfminer.high_level.extract_text(r'C:\Users\Jay\Desktop\Python\Auto Run File Builder\targa_kbz.pdf', run_file_laparams)

print(text)
