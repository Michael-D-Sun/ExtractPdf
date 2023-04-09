import os

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
# from pdfminer.layout import LAParams
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage, LTChar
from pdfminer.converter import PDFPageAggregator
from openpyxl import Workbook
from PIL import Image
import pytesseract
import io


def extract_image_from_pdf(file_name):
    # Set parameters for analysis.
    laparams = LAParams()
    rsrcmgr = PDFResourceManager()

    # Create a PDF page aggregator object.
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Open a PDF file.
    fp = open("{0}.pdf".format(file_name), 'rb')

    # Create a PDF parser object associated with the file object.
    parser = PDFParser(fp)
    # Create a PDF document object that stores the document structure.
    # Supply the password for initialization.
    document = PDFDocument(parser)

    for p, page in enumerate(PDFPage.create_pages(document)):
        interpreter.process_page(page)
        # receive the LTPage object for the page.
        layout = device.get_result()

        for node in layout:
            # print(type(node))

            if isinstance(node, LTFigure):
                for n in node:
                    # print(" {}".format(type(n)))
                    if n.stream:
                        buffer = io.BytesIO(n.stream.get_data())
                        # with open("test_data_{}.img".format(p),'wb') as fp :
                        #    fp.write(buffer.read())
                        pillow_object = Image.open(buffer)
                        pillow_object.save("test_image_{}.jpg".format(p))


def handle_invoice(text, sheet):
    for temp in text.split('\n'):
        row = None
        if len(temp) >= 5:
            row = temp.split()[-4::]
            row.insert(0, temp.split()[0])
        else:
            row = temp.split()
        sheet.append(row)


g_start = False
g_max_num = 0


def handle_packaging_list(text, sheet):
    row = None
    global g_start
    global g_max_num
    for temp in text.split('\n'):
        if ("CARTON" in temp) or ("PALLET" in temp) and (not g_start):
            g_start = True
            elements = temp.split()
            g_max_num = len(elements)
            # print(temp)
            # print("g_max_num: {0}".format(g_max_num))

        # handle   @xx @xx @xx
        if g_start and (len(temp.split()) < g_max_num - 2) and ("@" in temp):
            row = temp.split()
            row.insert(0, None)
            row.insert(0, None)
            row.insert(0, None)

        # handle 123 ' ' ' ' 123 123
        elif g_start and (len(temp.split()) < g_max_num - 2) and ("@" not in temp):
            row = temp.split()
            row.insert(0, None)
            row.insert(0, None)
            row.insert(0, None)
            row.insert(4, None)
            row.insert(4, None)
        else:
            # handle normal line
            row = temp.split()

        # print(row)
        sheet.append(row)


def handle_other_work(text, sheet):
    pass


def extract_text_and_and_write_to_excel(file_name):
    # extract text
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    cwd = os.getcwd()
    page_count = 0

    # delete existed .xlsx
    for path in os.listdir(cwd):
        if ".xlsx" in path:
            os.remove(path)

    work_book = Workbook()
    sheet = None
    process_work = None
    title = None

    for path in os.listdir(cwd):
        if ".jpg" in path:
            file = Image.open(path)
            config = r'-c tessedit_char_blacklist=  --psm 6'
            text = pytesseract.image_to_string(file, config=config)

            print(text)

            # if text.startswith("INVOICE"):
            #     print("Start to handle invoice pages.")
            #     process_work = "invoice"
            #     title = "Page0"
            #
            # if text.startswith("PACKING LIST"):
            #     process_work = "packing_list"
            #     print("Finish handling invoice pages.")
            #     print("Start to handle packing list pages.")
            #     title = "Page1"

            if "SHENGGAO" in text:
                page_count += 1
                title = "Page{0}".format(page_count)
                process_work = "work{0}".format(page_count)

            # create sheet
            if title not in work_book.sheetnames:
                sheet = work_book.create_sheet(title)

            # process image text
            match process_work:
                case "work1":
                    handle_invoice(text, sheet)

                case "work2":
                    handle_packaging_list(text, sheet)

                case _:
                    handle_other_work(text, sheet)

    print("Finish handling packing list pages.")
    del work_book["Sheet"]
    work_book.save("{0}.xlsx".format(file_name))
    print("All the data has been save to {0}.xlsx".format(file_name))


if __name__ == "__main__":
    print("*******************ExtractPdf Tool**********************")
    print("*Description : Extract Pdf data, output data to excel  *")
    print("*Usage       : Input the pdf name                      *")
    print("********************************************************")
    file_name = input("Please enter pdf file name: ")
    extract_image_from_pdf(file_name)
    extract_text_and_and_write_to_excel(file_name)