import tkinter as tk
from tkinter import messagebox as MB
from tkinter import scrolledtext
from tkinter import scrolledtext as sc
import tkinter.ttk as ttk1
from tkinter import *
import math
'''import base64
# import netmiko
import time
import threading
import xlsxwriter
import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border, Alignment, Side
from openpyxl.styles import colors
from openpyxl.cell import Cell
from openpyxl.styles import Font
from openpyxl.styles import colors
from openpyxl.cell import Cell
import pandas
import xlwt
from xlutils.copy import copy
from datetime import datetime
import pyexcel
# for mail
import email, smtplib, ssl
import cisco
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText'''

#------------------------------------------------------------------
#                          Scrollbar
#------------------------------------------------------------------
class Scrollable(tk.Frame):
    """
       Make a frame scrollable with scrollbar on the right.
       After adding or removing widgets to the scrollable frame,
       call the update() method to refresh the scrollable area.
    """
    def __init__(self, frame, width=10):
        scrollbar = tk.Scrollbar(frame, width=width)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=False)

        self.canvas = tk.Canvas(frame, width=1490, height=530, yscrollcommand=scrollbar.set)
        self.canvas.pack()
        scrollbar.config(command=self.canvas.yview)
        self.canvas.bind('<Configure>', self.__fill_canvas)
        # base class initialization
        tk.Frame.__init__(self, frame)
        # assign this obj (the inner frame) to the windows item of the canvas
        self.windows_item = self.canvas.create_window(0, 0, window=self, anchor=tk.NW)

    def __fill_canvas(self, event):
        "Enlarge the windows item to the canvas width"
        canvas_width = event.width
        self.canvas.itemconfig(self.windows_item, width=canvas_width)

    def update(self):
        "Update the canvas and the scrollregion"
        self.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox(self.windows_item))

#------------------------------------------------------------------
#                          Scrollbar
#------------------------------------------------------------------
root = tk.Tk()
# root.wm_iconbitmap('amex_test_icon_I6D_icon.ico')
root.title("CAMPUS BOM AND IP REQUEST PORTAL LOGIN")
root.state("zoom")
root_background_image = tk.PhotoImage(file='bg2.png')
# root_background_image = root_background_image.zoom(2, 2)
Main_page_logo = tk.PhotoImage(file="Amex new logo.gif")
w = root_background_image.width()
h = root_background_image.height()
root.geometry("%dx%d" % (w, h))
background_label = tk.Label(root, image=root_background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

font = ('times', 20, 'bold')


def charlimit3(entry_text):
    if len(entry_text.get()) > 0:
        entry_text.set(entry_text.get()[:3])

def charlimit5(entry_text_building):
    if len(entry_text_building.get()) > 0:
        entry_text_building.set(entry_text_building.get()[:5])

def IP_Frame1(frame=tk.Frame(root)):
    frame.destroy()
    Cancel_btn.destroy()
    Go1_btn.destroy()
    floor_leaf_info = {}
    leaf_per_floor_list = []
    frame1 = tk.Frame(root, bd=6, relief=tk.SUNKEN, bg='black')
    frame1.place(relx=0.5, rely=0.15, relwidth=0.6, relheight=0.2, anchor='center')
    label = tk.Label(frame1, text="Enter City or \nAirport Code(3 Char)\n", font=font, fg='white',bg='black')
    label.grid(row=0, column=1)
    entry_text = StringVar()
    city = tk.Entry(frame1, width=25, textvariable=entry_text)
    city.grid(row=0, column=2)
    entry_text.trace("w", lambda *args: charlimit3(entry_text))
    entry_text_building = StringVar()
    label = tk.Label(frame1, text=" Enter The Building Name\n(2-5 Char)\n ", font=font, fg='white',bg='black')
    label.grid(row=0, column=4)
    building_name = tk.Entry(frame1, width=25, textvariable=entry_text_building)
    building_name.grid(row=0, column=5)
    entry_text_building.trace("w", lambda *args: charlimit5(entry_text_building))
    label = tk.Label(frame1, text="Select Region         ", font=font, fg='white',bg='black')
    label.grid(row=1, column=1)
    region = ttk1.Combobox(frame1, width=23)
    region['values'] = ('Select ', 'USA/CAN', 'LAC', 'JAPA', 'EMEA')
    region.current(0)
    region.grid(row=1, column=2)
    lable = tk.Label(frame1, text="Total Number of Floor", font=font, fg='white',bg='black')
    lable.grid(row=1, column=4)
    total_floor_entry = tk.Entry(frame1, width=25)
    total_floor_entry.grid(row=1, column=5)

    def IP_Frame2():
        floor_number_list=[]
        floor_number_list_values = []
        frame2 = tk.Frame(root, bd=6, relief=tk.SUNKEN, bg='black')
        frame2.place(relx=0.5, rely=0.4, relwidth=0.6, anchor='center')
        total_floor = int(total_floor_entry.get())
        #frame1.destroy()
        last_row_for_button=0
        for i in range(1, total_floor + 1):
            print(i)
            lable = tk.Label(frame2, text=" S.No.", font=font, fg='white', bg='black')
            lable.grid(row=1, column=1)
            lable = tk.Label(frame2, text=str(i), font=font, fg='white', bg='black')
            lable.grid(row=1 + i, column=1)

            lable = tk.Label(frame2, text=" Enter the Number of Floor", font=font, fg='white',bg='black')
            lable.grid(row=1 + i, column=2)
            flr_num_entry = tk.Entry(frame2, width=20)
            flr_num_entry.grid(row=1 + i, column=3)
            floor_number_list.append(flr_num_entry)
            print(floor_number_list)
            lable = tk.Label(frame2, text=" Enter Number of Leaf per floor", font=font, fg='white',bg='black')
            lable.grid(row=1 + i, column=5)
            leaf_per_floor_entry = tk.Entry(frame2, width=20)
            leaf_per_floor_entry.grid(row=1 + i, column=6)
            leaf_per_floor_list.append(leaf_per_floor_entry)
            last_row_for_button=i


        def IP_Frame3():
            for i in range(total_floor):
                floor_number_list_values.append(floor_number_list[i].get())
            print(floor_number_list_values)
            floor_number_list_intvalues = [int(i) for i in floor_number_list_values]
            count = False
            for i1 in floor_number_list_intvalues:
                if floor_number_list_intvalues.count(i1) >= 2:
                    count = True
                    print("Count:", count)
                    break
            if count == True:
                res = MB.askretrycancel('Message title', 'Floor name same')
                if res == True:
                    print('check')
                    IP_Frame2()
            else:

                frame3 = tk.Frame(root, bd=6, relief=tk.SUNKEN, height=100, bg='black')
                frame3.place(relx=0.5, rely=0.3, anchor='n')
                sum=0
                for i in range(len(floor_number_list)):
                    floor_leaf_info[floor_number_list[i].get()]=int(leaf_per_floor_list[i].get())
                    sum=sum+int(leaf_per_floor_list[i].get())
                print(floor_leaf_info)
                #flr = int(total_floor_entry.get())
                scrollable_body = Scrollable(frame3, width=10)
                flr_num_txt_list = []
                device_combo_list = []
                device_floor_loc_list = []
                device_vendor_list = []
                device_type_list = []
                device_pos_list = []
                device_model_list = []
                i=0
                #frame2.destroy()
                for j in floor_leaf_info:

                    for i1 in range(floor_leaf_info[j]):
                        i=i+1
                        label = tk.Label(scrollable_body, text="Floor \nNumber", font=font)
                        label.grid(row=10, column=1)
                        label1 = tk.Label(scrollable_body, text=j, font=font, width=10)
                        label1.grid(row=10 + i, column=1)

                        l1 = tk.Label(scrollable_body, text="Enter Serial \nNumber \n", font=font)
                        l1.grid(row=10, column=2)

                        flr_num_txt = tk.Entry(scrollable_body, width=18)  # list
                        flr_num_txt.grid(row=10 + i, column=2)
                        flr_num_txt_list.append(flr_num_txt)

                        device_use = tk.Label(scrollable_body, text="    Select Device \nUse \n", font=font)
                        device_use.grid(row=10, column=3)
                        device_combo = ttk1.Combobox(scrollable_body, width=20)  # list
                        device_combo['values'] = ('Select ', 'Campus', 'Management')
                        device_combo.current(0)
                        device_combo.grid(row=10 + i, column=3)
                        device_combo_list.append(device_combo)

                        device_vendor = tk.Label(scrollable_body, text="Select the \n Device Vendor \n", font=font)
                        device_vendor.grid(row=10, column=5)
                        device_vendor_combo = ttk1.Combobox(scrollable_body, width=20)  # list
                        device_vendor_combo['values'] = ('Select ', 'Cisco', 'Cumulus', 'Lantronix', 'Arista', 'Netflow')
                        device_vendor_combo.current(0)
                        device_vendor_combo.grid(row=10 + i, column=5)
                        device_vendor_list.append(device_vendor_combo)

                        device_type = tk.Label(scrollable_body, text="    Select Device \nType\n", font=font)
                        device_type.grid(row=10, column=6)
                        device_type_combo = ttk1.Combobox(scrollable_body, width=20)  # list
                        device_type_combo['values'] = (
                            'Select ', 'Super Spine', 'Spine', 'Core Services', 'Leaf', 'Terminal Server', 'Netflow', 'Layer 2')
                        device_type_combo.current(0)
                        device_type_combo.grid(row=10 + i, column=6)
                        device_type_list.append(device_type_combo)

                        device_model = tk.Label(scrollable_body, text="    Select Device \nModel\n", font=font)
                        device_model.grid(row=10, column=7)
                        device_model_combo = ttk1.Combobox(scrollable_body, width=20)  # list
                        device_model_combo['values'] = ('Select ', '4610-54P-O-AC-F-US', '210-ADUX')
                        device_model_combo.current(0)
                        device_model_combo.grid(row=10 + i, column=7)
                        device_model_list.append(device_model_combo)

                        device_num3 = tk.Label(scrollable_body, text="Enter Device \nPosition \n (2 Char)", font=font)
                        device_num3.grid(row=10, column=8)
                        device_pos = tk.Entry(scrollable_body, width=20)  # list
                        device_pos.grid(row=10 + i, column=8)
                        device_pos_list.append(device_pos)
                scrollable_body.update()

            def Config_Function():
                serialvar=0
                device_combo_var=0
                hostname_list=[]
                device_type_var_list=[]
    
                #for j in device_combo_list:
                 #   device_combo_var=j.get()
                  #  if device_combo_var=="Campus":
                   #     device_combo_var="Cam"
                    #elif device_combo_var=="Management":
                     #   device_combo_var="Mgt"
                    #print(device_combo_var)
                #--------------------------------------------
                #            CREATE EXCEL File
                # --------------------------------------------
                now = datetime.now()
                runtime = 'Config File{}.xlsx'.format(now.strftime("%c")).replace(':', '_')
                filename = 'TESTFILE_' + runtime + '.xlsx'  # Creating a file name and injecting the runtime variable within the string
                wb = xlsxwriter.Workbook(filename)  # Assigning the file name to the workbook and creating the excel workbook
                ws1 = wb.add_worksheet('Results')  # creating a new sheet within the excel workbook (REQUIRED)
                headerFormat = wb.add_format(
                {'bold': True,  # Creating Dictionary for header format within the "add_format" function
                 'font_name': 'Calibri',
                 'font_size': 14,
                 'bg_color': 'gray',
                 'align': 'center',
                 'valign': 'vcenter',
                 'text_wrap': True})
                normalFormat = wb.add_format(
                {'font_name': 'Calibri',  # Creating Dictionary for header format within the "add_format" function
                 'font_size': 12,
                'valign': 'vcenter',
                'text_wrap': True})
                row=1
                ws1.set_column(0, 0, 15,normalFormat)  # Set the column with 0 = Column A in Excel or Row 1 in excel (Range of column)
                ws1.set_column(1, 1, 35, normalFormat)
                ws1.set_column(2, 2, 25, normalFormat)  # Altered in SESSION2
                ws1.set_column(3, 3, 30, normalFormat)  # Altered in SESSION2
                ws1.set_column(4, 4, 25, normalFormat)  # Altered in SESSION2
                ws1.set_column(5, 5, 25, normalFormat)  # Altered in SESSION2
    
    
                ws1.write(0,0,'S No.',headerFormat)
                ws1.write(0, 1, 'Device Serial Number',headerFormat)  # Setting Headers (0, 0, = row, column in spreadsheet)
                ws1.write(0, 2, 'Hostname', headerFormat)
                ws1.write(0, 3, 'Region', headerFormat)  # Altered in SESSION2
                ws1.write(0, 4, 'Device Model', headerFormat)  # Altered in SESSION2
                ws1.write(0, 5, 'Device Role', headerFormat)  # Altered in SESSION2
    
                for i in range(len(flr_num_txt_list)):
                    cNo=country_txt.get()
                    bName=building_name.get()
                    device_combo_var = device_combo_list[i].get()
                    if device_combo_var == "Campus":
                        device_combo_var = "cam"
                    elif device_combo_var == "Management":
                        device_combo_var = "Mgt"
                    device_flr_var=device_floor_loc_list[i].get()
                    device_vendor_var = device_vendor_list[i].get()
                    if device_vendor_var == "Cisco":
                        device_vendor_var = "ci"
                    elif device_vendor_var == "Cumulus":
                        device_vendor_var = "cn"
                    elif device_vendor_var == "Lantronix":
                        device_vendor_var = "lx"
                    elif device_vendor_var == "Arista":
                        device_vendor_var = "as"
                    device_type_var = device_type_list[i].get()
                    if device_type_var == "Super Spine":
                        device_type_var = "ss"
                    elif device_type_var == "Spine":
                        device_type_var = "sp"
                    elif device_type_var == "Core Services":
                        device_type_var = "sv"
                    elif device_type_var == "Leaf":
                        device_type_var = "lf"
                    elif device_type_var == "Terminal Server":
                        device_type_var = "ts"
                    elif device_type_var == "Netflow":
                        device_type_var = "nf"
                    elif device_type_var == "Layer 2":
                        device_type_var = "sw"
                    #device_type_var_list.append(device_type_var)
                    device_position_var = device_pos_list[i].get()
    
                    hostname_list.append(str(cNo)+"-"+str(bName)+str(device_combo_var)+str(device_flr_var)+str(device_vendor_var)+str(device_type_var)+str(device_position_var))
    
    
                for i in range(len(flr_num_txt_list)):
                    serialvar = flr_num_txt_list[i].get()
                    print(serialvar)
                    ws1.write(row,0,i+1)
                    ws1.write(row, 1, serialvar)
                    ws1.write(row,2,hostname_list[i])
                    ws1.write(row,3,region_combo.get())
                    ws1.write(row, 4, device_model_combo.get())
                    ws1.write(row, 5, device_type_list[i].get())
    
                    row += 1
                wb.close()

        Cancel_btn4 = tk.Button(root, text="BACK", command=exit)
        Cancel_btn4.place(x=0.5,y=0.9,relx=0.6,rely=0.9, anchor='n')
        Go1_btn4 = tk.Button(root, text='Export', command=exit)
        Go1_btn4.place(x=0.5,y=0.9,relx=0.4,rely=0.9, anchor='n')

        Cancel3_btn = tk.Button(frame2, text="BACK1", command=lambda:IP_Frame1(frame2))
        Cancel3_btn.grid(row=10+last_row_for_button,column=5)
        Go3_btn = tk.Button(frame2, text='NEXT', command=IP_Frame3)
        Go3_btn.grid(row=10+last_row_for_button,column=3)

    Cancel2_btn = tk.Button(frame1, text="BACK", command=exit)
    Cancel2_btn.place(relx=0.7, rely=0.8, anchor='center')
    Go2_btn = tk.Button(frame1, text='NEXT', command=IP_Frame2)
    Go2_btn.place(relx=0.4, rely=0.8, anchor='center')


Cancel_btn = tk.Button(root, text="BACK", command=exit)
Cancel_btn.grid(row=10 + 7 + 1, column=2)
Go1_btn = tk.Button(root, text='NEXT115', command=IP_Frame1)
Go1_btn.grid(row=10 + 7 + 1, column=3)
root.mainloop()

