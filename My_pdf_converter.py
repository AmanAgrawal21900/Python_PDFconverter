# Importing Modules
import os
from fpdf import *
from docx2pdf import convert
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog
from emoji import emojize

# Making global variables and dictionaries
select = 0
values = {"Docx to PDF": "1", "Txt to PDF": "2"}
file_path = ''
path = ''


# Making Functions
def change():
    global select
    if select == 1:
        do_to_pdf()
    if select == 2:
        text_to_pdf()


def selection():
    global select
    select = v.get()


def browse():
    global file_entry, select, file_path
    file_entry.delete("0", END)
    file_path = filedialog.askopenfilename(
        filetype=[("All Files", "*.*"), ("Docx Files", "*.docx"), ("Text Files", "*.txt")]
    )
    if not file_path:
        return None
    file_entry.insert('1', file_path)


def do_to_pdf():
    global file_entry, file_path, path
    path = simpledialog.askstring("My_PDF_converter", "Enter name of pdf")
    if path == '':
        base_name = os.path.basename(file_entry.get())
        filename = os.path.splitext(base_name)[0]
        convert(file_entry.get(), 'C:/Users/Aman Agrawal/Desktop/' + filename + '.pdf')
        messagebox.showinfo("My_PDF_converter", filename + ".pdf" + " Created Successfully")
    elif '.pdf' in path:
        convert(file_entry.get(), 'C:/Users/Aman Agrawal/Desktop/' + path)
        messagebox.showinfo("My_PDF_converter", path + " Created Successfully")
    else:
        pdf_name = str(path) + '.pdf'
        convert(file_entry.get(), 'C:/Users/Aman Agrawal/Desktop/' + pdf_name)
        messagebox.showinfo("My_PDF_converter", pdf_name + " Created Successfully")


def text_to_pdf():
    global file_entry, path, file_path
    # save FPDF() class into a variable pdf
    pdf = FPDF()
    # Add a page
    pdf.add_page()
    # set style and size of font that you want in the pdf
    pdf.set_font("Arial", size=10)
    # open the text file in read mode

    f = open(file_entry.get(), "r", encoding='utf-8')
    # insert the texts in pdf
    for x in f:
        """using latin-1 encoding and decoding MUST"""
        text2 = x.encode('latin-1', 'replace').decode('latin-1')
        pdf.cell(200, 5, txt=text2, ln=1, align='L')

    path = simpledialog.askstring('My_PDF_convertor', 'Enter name of your PDF')
    if path == '':
        base_name = os.path.basename(file_entry.get())
        filename = os.path.splitext(base_name)[0] + '.pdf'
        pdf.output('C:/Users/Aman Agrawal/Desktop/' + filename)
        messagebox.showinfo("My_PDF_converter", filename + ".pdf" + " Created Successfully")
    elif '.pdf' in path:
        pdf.output('C:/Users/Aman Agrawal/Desktop/' + path)
        messagebox.showinfo("My_PDF_converter", path + " Created Successfully")
    else:
        pdf_name = str(path) + '.pdf'
        pdf.output('C:/Users/Aman Agrawal/Desktop/' + pdf_name)
        messagebox.showinfo("My_PDF_convertor", pdf_name + " Created Successfully")


# Making Window
win = Tk()
win.title("My_PDF_Converter")
win.geometry("490x225+400+100")
win.config(bg="white")
win.resizable(0, 0)
v = IntVar()
# Defining Functions

# Making Title frame
title_frame = Frame(master=win, bg="yellow", height=30)
title_frame.pack(fill=X, side=TOP)
title_label = Label(master=title_frame, text="PDF Converter -- Convert .docx and .txt to PDF",
                    font=("Verdana", 12, "italic"), bg="khaki")
title_label.pack()

# Making file_name frame
file_frame = Frame(master=win, bg="medium spring green", height=100)
file_frame.pack(side=TOP, fill=X)
file_label = Label(master=file_frame, text="Enter file name : ", font=("Verdana", 11),
                   anchor=W, width=75, bg='medium spring green',
                   pady=5)
file_label.pack()
file_entry = Entry(master=file_frame, bg="white", width=60)
file_entry.pack()

# Making button frame
btn_frame = Frame(master=win, bg="medium spring green", pady=10, padx=10)
btn_frame.pack(side=TOP, fill=X)
browse_btn = Button(master=btn_frame, text="Browse", height=1, width=10, bg="lime", command=browse)
browse_btn.pack(side=LEFT, padx=10)
convert_btn = Button(master=btn_frame, text="Convert !" + emojize(':thumbs_up:'), height=1,
                     width=15, bg="lime", command=change)
convert_btn.pack(side=LEFT, padx=10)

for text, values in values.items():
    Radiobutton(master=btn_frame, text=text, value=values, font=("Verdana", 11, "italic"), variable=v,
                bg="medium spring green", justify="left",
                command=selection).pack(side="top", padx=5, pady=5, anchor=W)

# Making output frame
out_frame = Frame(master=win, bg="forest green", pady=10, padx=10)
out_frame.pack(side="top")
out_label = Label(master=out_frame, text="", font=("Verdana", 11),
                  anchor=W, width=75, bg='forest green',
                  pady=5)
out_label.pack(side="top")
# Making mainloop
win.mainloop()
