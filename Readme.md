1. Prerequisites  
Install python3  
pip install pdfminer  
pip install pytesseract  
pip install pyopenxl   
pip install pyinstaller 
Install tesseract from [tesseract](https://github.com/UB-Mannheim/tesseract/wiki)  

2. Pack it as exe file  
pyinstaller -F --noconsole --version-file file_version_info.txt ExtractPdf.py 

3. How to use  
Double click ExtractPdf.exe  
Click Start button to start converting PDF to Excel
Get the Excel file in the same folder with ExtractPDF.exe

4. How to add new function