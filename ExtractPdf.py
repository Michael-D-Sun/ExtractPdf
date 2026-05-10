import os
import io
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage, LTChar
from pdfminer.converter import PDFPageAggregator
from openpyxl import Workbook
from PIL import Image
import pytesseract

g_text = None


def extract_images_from_pdf(file_path, output_dir=None):
    """Extract all pictures/images from a PDF and save them to a dedicated folder.

    Args:
        file_path: Path to the PDF file.
        output_dir: Directory to save extracted images. Defaults to
                    '<pdf_name>_images' in the same directory as the PDF.

    Returns:
        List of saved image file paths.
    """
    if output_dir is None:
        base_name = os.path.splitext(file_path)[0]
        output_dir = f"{base_name}_images"
    os.makedirs(output_dir, exist_ok=True)

    laparams = LAParams()
    rsrcmgr = PDFResourceManager()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    saved_images = []
    with open(file_path, 'rb') as fp:
        parser = PDFParser(fp)
        document = PDFDocument(parser)

        image_index = 0
        for page_num, page in enumerate(PDFPage.create_pages(document)):
            interpreter.process_page(page)
            layout = device.get_result()

            for element in layout:
                if isinstance(element, LTFigure):
                    for item in element:
                        if isinstance(item, LTImage) and item.stream:
                            try:
                                buffer = io.BytesIO(item.stream.get_data())
                                img = Image.open(buffer)
                                fmt = img.format.lower() if img.format else 'png'
                                name = f"page{page_num + 1}_img{image_index + 1}.{fmt}"
                                save_path = os.path.join(output_dir, name)
                                img.save(save_path)
                                saved_images.append(save_path)
                                image_index += 1
                            except Exception as e:
                                print(f"Warning: could not extract image on page {page_num + 1}: {e}")

    if saved_images:
        print(f"Extracted {len(saved_images)} images to {output_dir}")
    else:
        print("No images found in the PDF.")
    return saved_images


def handle_invoice(text, sheet):
    for temp in text.split('\n'):
        row = temp.split()
        while (len(row) > 3) and (not row[2][0].isdigit()):
            del row[2]
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

        # handle xx space space xx xx
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
    print("Currently, don't support such kind of pdf format.")


def extract_text_and_and_write_to_excel(file_name, image_dir):
    # extract text
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    work_count = 0
    works = ["invoice", "packing_list"]

    # delete existed .xlsx
    for path in os.listdir(os.getcwd()):
        if ".xlsx" in path:
            os.remove(path)

    work_book = Workbook()
    process_work = None
    sheet = None

    for path in os.listdir(image_dir):
        if path.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.tiff')):
            file = Image.open(os.path.join(image_dir, path))
            config = r'-c tessedit_char_blacklist=  --psm 6'
            content = pytesseract.image_to_string(file, config=config)
            # print(context)

            if "SHENGGAO" in content:
                if 2 <= work_count:
                    break
                else:
                    print("work_count:{0}".format(work_count))
                    process_work = works[work_count]
                    work_count += 1

            # create sheet
            if process_work not in work_book.sheetnames:
                sheet = work_book.create_sheet(process_work)

            # process image context
            match process_work:
                case "invoice":
                    handle_invoice(content, sheet)

                case "packing_list":
                    handle_packaging_list(content, sheet)

                case _:
                    handle_other_work(content, sheet)

    print("Finish handling packing list pages.")
    if "Sheet" in work_book.sheetnames:
        del work_book["Sheet"]
    work_book.save("{0}.xlsx".format(file_name.split(".")[0]))
    print("All the data has been save to {0}.xlsx".format(file_name.split(".")[0]))


def open_file():
    try:
        # open file and return full path
        path = filedialog.askopenfilename()
        file_path.set(path)
    except Exception as e:
        print("You haven't selected any file yet", e)
        file_path.set("")
    print("filename is {0}".format(file_path.get()))
    text.delete(1.0, END)
    start_button.config(bg='SystemButtonFace')


def extract_pdf():
    if file_path.get() == "" or ".pdf" not in file_path.get():
        messagebox.showerror("Error", "Please select PDF file")
    else:
        (directory, file_name) = os.path.split(file_path.get())
        print("directory:{0}, file_name:{1}, file_path:{2}".format(directory, file_name, file_path.get()))
        image_dir = extract_images_from_pdf(file_path.get())
        extract_text_and_and_write_to_excel(file_name, image_dir)
        info = "Finish converting PDF to Excel. Excel is stored at {0}".format(directory+"/"+file_name.split(".")[0]
                                                                               +".xlsx")
        text.delete(1.0, END)
        text.insert(INSERT, info)
    start_button.config(bg='green')


def extract_pictures_only():
    if file_path.get() == "" or ".pdf" not in file_path.get():
        messagebox.showerror("Error", "Please select PDF file")
    else:
        saved = extract_images_from_pdf(file_path.get())
        if saved:
            count = len(saved)
            output = os.path.dirname(saved[0])
            info = f"Extracted {count} image(s) to {output}"
            text.delete(1.0, END)
            text.insert(INSERT, info)
            extract_pictures_button.config(bg='green')
        else:
            messagebox.showinfo("Info", "No images found in the PDF.")


if __name__ == "__main__":
    root_window = Tk()
    root_window.geometry('500x350')
    root_window.title("ExtractPdf")
    file_path = tk.StringVar()

    entry = tk.Entry(root_window, width=40, textvariable=file_path)
    entry.grid(row=0, column=1, padx=5, pady=5)

    text = tk.Text(root_window, name="text", width=40, height=10, padx=5, pady=5, undo=True, autoseparators=False)
    text.grid(row=3, column=1)

    browse_button = tk.Button(root_window, text='BrowserFile', padx=5, pady=5, command=open_file)
    browse_button.grid(row=0, column=2)

    start_button = tk.Button(root_window, width=40, text='Convert PDF to Excel', padx=5, pady=5, command=extract_pdf)
    start_button.grid(row=1, columnspan=2)

    extract_pictures_button = tk.Button(root_window, width=40, text='Extract Pictures', padx=5, pady=5, command=extract_pictures_only)
    extract_pictures_button.grid(row=2, columnspan=2)

    root_window.mainloop()
