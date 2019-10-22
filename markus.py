import csv
import random
import string
import os
import configparser
import codecs
import win32com.client as win32
from docxtpl import DocxTemplate, Listing
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory


class DocumentTemplate:
    file = ""  # Template file location
    name = ""  # Name of template
    naming = ""  # Naming template for resulting files


main_config = configparser.RawConfigParser()
main_config.read_file(codecs.open("markus.conf", "r", "utf8"))

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

raw_var = main_config["current"]["var_to_indicate_row"]

messagebox.showinfo("Warning!", "Please close all Excel documents or you can lost unsaved data!!!")
xlsm_input = askopenfilename(initialdir=main_config["current"]["source_file_dir"], title="Choose data source Excel file",
                             filetypes=(("Excel Document", "*.xl*"),))  # show an "Open" dialog box and return the path to the selected file

templates = []  # list of tamplates

templates_coutn = int(main_config["current"]["docx_templates_count"])
for i in range(1, templates_coutn + 1):
    dt = DocumentTemplate()
    dt.name = main_config["current"]["docx_template_name_" + str(i)]
    dt.file = askopenfilename(initialdir=main_config.get("current", "docx_template_dir_" + str(i), fallback=""), title="Please select template for " + dt.name,
                              filetypes=(("Word Document", "*.docx"),))
    dt.naming = main_config["current"]["docx_template_naming_" + str(i)]
    templates.append(dt)

save_dir = askdirectory(title="Choose directory to save files")

excel = win32.Dispatch("Excel.Application")

excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(xlsm_input)

random_name = ''.join(random.choices(string.ascii_lowercase + string.digits, k=4))
data_file = os.environ["temp"] + "\\" + random_name + ".csv"

wb.Sheets(main_config["current"]["complete_data_sheet_name"]).Select()

wb.SaveAs(data_file, 23)

wb.Close()
excel.Quit()
try:
    with open(data_file, "r") as csv_file:
        dialect = csv.Sniffer().sniff(csv_file.readline())
        csv_file.seek(0)
        csv_reader = csv.reader(csv_file, dialect)

        rows = list()
        for row in csv_reader:
            rows.append(row)

        doc_vars = rows[0]
        raw_var_index = 0
        for i in range(0, len(doc_vars)):
            if doc_vars[i] == raw_var:
                raw_var_index = i
                break

        first_data_row = 1 + int(main_config["current"]["lines_to_skip"])
        for i in range(first_data_row, len(rows)):
            if str(rows[i][raw_var_index]) == "":
                continue
            data = dict()
            for l in range(0, len(rows[0])):
                if doc_vars[l] == '':
                    continue
                if "\n" in str(rows[i][l]):
                    data[str(doc_vars[l])] = Listing(str(rows[i][l]))  # Listing is used to send multiline strings
                else:
                    data[str(doc_vars[l])] = str(rows[i][l])

            for t in templates:
                doc = DocxTemplate(t.file)
                doc.render(data)
                doc.save(save_dir + "/" + t.naming % data)
except Exception as err:
    print(err)
finally:
    if os.path.exists(data_file):
        os.remove(data_file)
messagebox.showinfo("Markus", "Done")