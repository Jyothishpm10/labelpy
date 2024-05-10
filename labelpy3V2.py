"""This program is intened to generate EPP label according to al rashideen order sheet. The order 
sheet excel file is edited in following way. Columns must be name,brand,w,h,one,two,[]...,nos,media,jobno,
category all other columns are removed. It is important to note that columns are in lower case and 
it can not be changed. Column names are explained below.

name : POS
brand : brand name like chesterfeild or malboro etc.
w : actual width
h : actual height
one, two ,... : represents the each artworks columns 
nos : total qty
media : which media is used like duratrance or polycarbonate
jobno : specific jobno
category : belongs to which category usualy dubai or auh other wise branded or unbranded like that

Blank cells under artwork columns are replaced by '' without space
Unwanted rows are removed and saved as csv file.
The result will create a folder with jobno and save the pdf file init. """

import re
import csv
import os
from fpdf import FPDF
import inflect
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog


class PDF(FPDF):
    """pdf class"""

    def label_body(
        self, name, category, media, width, height, qty, total_pack, art_col, jobno, cust_name, brand
    ):
        """pdf body"""
        self.add_page()
        self.set_font("helvetica", "B", 14)
        self.set_margins(5, 25, -5)
        self.cell(
            132,
            30,
            cust_name,
            border=0,
            new_x="LMARGIN",
            new_y="NEXT",
        )
        self.multi_cell(
            132,
            20,
            f"{name} - {category} - {media} - {brand}",
            border=0,
            max_line_height=15,
            new_x="LMARGIN",
            new_y="NEXT",
        )
        self.cell(
            132,
            25,
            f"Actual size   W:{width}cm  H :{height}cm",
            border=0,
            new_x="RIGHT",
            new_y="TMARGIN",
        )
        self.set_font("helvetica", "B", 30)
        self.cell(68, 75, "DIP", border=0, align="C", new_x="LMARGIN", new_y="NEXT")
        self.set_font("helvetica", "B", 14)
        self.multi_cell(
            93,
            25,
            f"{qty:35}{art_col}/{total_pack}",
            border=0,
            new_x="RIGHT",
            new_y="LAST",
        )
        self.cell(39, 25, qty, border=0, align="C")
        self.cell(40, 25, "", border=0)
        self.cell(28, 25, jobno, border=0, align="C")


def print_label(
    name, category, media, width, height, qty, col_num, total_pack, art_col, jobno, cust_name, brand
):
    """function to print label"""
    pdf = PDF("L", "mm", "A5")
    pdf.label_body(
        name, category, media, width, height, qty, total_pack, art_col, jobno, cust_name, brand
    )
    newname = re.sub(r"\W+", "",name)
    pdf.output(f"label_{newname}_w_{width}_h_{height}_art{col_num}_{jobno}.pdf")

def read_csv(file_name):
    data = []
    with open(file_name,'r',encoding='utf-8') as file:
        reader = csv.reader(file)
        line_list = [[value.strip() for value in line] for line in reader]
    keys = line_list.pop(0)
    for line in line_list:
            data.append({keys[i]:line[i] for i in range(len(keys))})
    return data, len(keys)

def write_label(dir, cust_name):
    '''the writing function'''
    merger = PdfMerger()
    engine = inflect.engine()
    data, total_columns = read_csv(dir)
    for row in data:
        value = list(row.values())
        total_pkt = total_columns - (value.count("") + 8)
        art_col = 1
        for col in range(1, (total_columns - 7)):
            col = engine.number_to_words(col)
            if row[col] == "":
                continue
            else:
                print_label(
                    row["name"].strip(),
                    row["category"].strip(),
                    row["media"].title(),
                    row["w"],
                    row["h"],
                    row[col],
                    col,
                    total_pkt,
                    art_col,
                    row["jobno"],
                    cust_name,
                    row["brand"]
                )
            art_col += 1

    pdf_list = [pdf for pdf in os.listdir() if pdf.endswith(".pdf")]
    pdf_name = os.path.basename(dir).removesuffix(".csv")
    for pdf in sorted(pdf_list):
        merger.append(pdf)
    merger.write(f"label_{pdf_name}.pdf")
    merger.close()
    os.renames(
        f"label_{pdf_name}.pdf",
        f"label_{pdf_name}/single_file/label_{pdf_name}.pdf",
    )
    os.chdir(".")
    for pdf in pdf_list:
        new_dir = os.path.join(f"label_{pdf_name}/seperate_files", pdf)
        os.renames(pdf, new_dir)


def main():

    window = tk.Tk()
    window.title('labelpy3')
    window.geometry('310x105')

    dir = tk.StringVar()
    name = tk.StringVar()
    
    def open_path():
        path = filedialog.askopenfilename()
        dir.set(path)


    cust_name_label = tk.Label(master=window,text='Customer Name:').grid(column=0,row=0,padx=5,pady=5,sticky=tk.NW)
    cust_name_entry= tk.Entry(master=window,width=32,textvariable = name).grid(column=1,row=0,padx=5,pady=5,columnspan=2)
    
    
    browse_button = tk.Button(master= window,text= 'browse',command=lambda:open_path()).grid(column=2,row=1,padx=5,pady=5)
    path_label = tk.Label(master=window,text='Path/:').grid(column=0,row=1,padx=5,pady=5,sticky=tk.E)
    path_entry = tk.Entry(master=window,textvariable=dir).grid(column=1,row=1,padx=5,pady=5)


    ok_button = tk.Button(master=window,text='OK',command  = lambda:write_label(dir.get(),name.get())).grid(column=0,row=2,padx=5,pady=5)
    cancel_button = tk.Button(master=window,text='Cancel', command =window.destroy).grid(column=2,row=2,padx=5,pady=5)



    window.mainloop()

if __name__ == "__main__":
    main()
