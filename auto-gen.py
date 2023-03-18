from openpyxl import load_workbook
import qrcode
from tkinter import *
import os 

try:
    dir = "./generated codes"
    os.mkdir(dir)
except FileExistsError:
    pass

def create_qr_code():
    info = []

    file_name =file_name_input.get()

    file = load_workbook(f"{file_name}.xlsx")
    sheet_use = file['Sheet1']

    def generate_qr_code(name, info):
        code = qrcode.make(info)
        code.save(f"{name}.png")

    for row in  range(2,300):
        scan_range = sheet_use[str(row)]
        for fields in scan_range:
            info.append(fields.value)

        ps_name = info[0]
        dpt_name = info[1]
        ps_company_id = info[2]
        data = f"\n\nName: {ps_name}\n\n Department: {dpt_name}\n\n Company Id: {ps_company_id}\n\n"

        generate_qr_code(ps_name, info=data)
        info.clear()

window = Tk()
window.title("Auto Qr code generator")
window.geometry("200x50")
window.config(bg="red")

file_name_label = Label(text="File Name")
file_name_label.grid(column=0, row=0)

file_name_input = Entry(text="File Name")
file_name_input.grid(column=1, row=0)

create_code = Button(text="Generate", bg="green", command=create_qr_code)
create_code.grid(column=1, row=1)

window.mainloop()


