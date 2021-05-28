import tkinter as tk
from tkinter import messagebox as MB
from tkinter import scrolledtext
from tkinter import scrolledtext as sc
import tkinter.ttk as ttk1
from tkinter import *
import math
import base64
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
from email.mime.text import MIMEText
class Scrollable(tk.Frame):
    """
       Make a frame scrollable with scrollbar on the right.
       After adding or removing widgets to the scrollable frame,
       call the update() method to refresh the scrollable area.
    """

    def __init__(self, frame, width=10):

        scrollbar = tk.Scrollbar(frame, width=width)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, expand=False)

        self.canvas = tk.Canvas(frame,width=1650, height=530,yscrollcommand=scrollbar.set)
        self.canvas.pack()
        scrollbar.config(command=self.canvas.yview)
        self.canvas.bind('<Configure>', self.__fill_canvas)
        # base class initialization
        tk.Frame.__init__(self, frame)
        # assign this obj (the inner frame) to the windows item of the canvas
        self.windows_item = self.canvas.create_window(0,0, window=self, anchor=tk.NW)

    def __fill_canvas(self, event):
        "Enlarge the windows item to the canvas width"

        canvas_width = event.width
        self.canvas.itemconfig(self.windows_item, width = canvas_width)

    def update(self):
        "Update the canvas and the scrollregion"

        self.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox(self.windows_item))

root = tk.Tk()
#root.wm_iconbitmap('amex_test_icon_I6D_icon.ico')
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
New_NGC_page_frame = tk.Frame(root, bd=6, relief=tk.SUNKEN)
New_NGC_page_frame.place(relx=0.5, rely=0.15, relwidth=0.8, relheight=0.25, anchor='center')
font = ('times', 20, 'bold')
entry_txt_list=[]

def charlimit5(entry_text_building):
    if len(entry_text_building.get()) >=0:
        entry_text_building.set(entry_text_building.get()[:5])
def charlimit3(entry_text):
    if len(entry_text.get())>=0:
        entry_text.set(entry_text.get()[:3])

def charlimit3for(entry_txt):
        if len(i.get())>0:
            i.set(i.get()[:3])
def New_NGC_Partial_Funtion():

    country = tk.Label(New_NGC_page_frame, text="Enter City Code or Airport Code\n(3 Char)", font=font)
    country.grid(row=0, column=1)
    entry_text=StringVar()
    country_txt = tk.Entry(New_NGC_page_frame, width=30,textvariable=entry_text)
    country_txt.grid(row=0, column=2)

    entry_text.trace("w", lambda *args: charlimit3(entry_text))


    region = tk.Label(New_NGC_page_frame, text="Select Region         ", font=font)
    region.grid(row=1, column=1)
    # region = tk.Entry(New_NGC_page_frame, width=30)
    # region.grid(row=1, column=2)
    region_combo = ttk1.Combobox(New_NGC_page_frame, width=27)
    region_combo['values'] = ('Select ', 'USA/CAN', 'LAC', 'JAPA', 'EMEA')
    region_combo.current(0)
    region_combo.grid(row=1, column=2)
    entry_text_building = StringVar()
    building_name = tk.Label(New_NGC_page_frame, text="     Enter The Building Name(2-5 Char) ", font=font)
    building_name.grid(row=0, column=4)
    building_name = tk.Entry( New_NGC_page_frame, width=30)
    building_name.grid(row=0, column=5)
    building_name.configure(textvariable=entry_text_building)
    entry_text_building.trace("w", lambda *args: charlimit5(entry_text_building))
    flr_lable = tk.Label(New_NGC_page_frame, text="          Total Number of Device", font=font)
    flr_lable.grid(row=2, column=4)
    flr_txt = tk.Entry(New_NGC_page_frame, width=30)
    flr_txt.grid(row=2, column=5)

    New_NGC_page_txt_frame111 = tk.Frame(root, bd=6, relief=tk.SUNKEN)
    New_NGC_page_txt_frame111.place(x=0.1, y=0.1, relx=0.5, rely=0.5, anchor='n')
    # New_NGC_page_labelfont111 = ('times', 20, 'bold')

    floor1_txt_frame = tk.Frame(root, bd=6, relief=tk.SUNKEN)
    floor1_txt_frame.place(relx=0.5, rely=0.2, anchor='n')

    floor2_txt_frame = tk.Frame(root, bd=6, relief=tk.SUNKEN)
    floor2_txt_frame.place(relx=0.5, rely=0.8, anchor='n')


    def NGC_CISCO_Function():
        NGC_frame_for_floor_input = tk.Frame(root, bd=6, relief=tk.SUNKEN,height=100)
        NGC_frame_for_floor_input.place(relx=0.5,rely=0.3,anchor='n')
        flr = int(flr_txt.get())
        scrollable_body=Scrollable(NGC_frame_for_floor_input,width=10)
        flr_num_txt_list=[]
        device_combo_list=[]
        device_floor_loc_list=[]
        device_vendor_list=[]
        device_type_list=[]
        device_pos_list=[]
        #entry_text_var_list=['none']

        for i in range(1, flr + 1):
            #entry_text_var_list.append('a'+str(i))
            l1 = tk.Label(scrollable_body, text="Enter Serial \nNumber \n", font=font)
            l1.grid(row=10, column=1)

            flr_num_txt = tk.Entry(scrollable_body, width=30)#list
            flr_num_txt.grid(row=10 + i, column=1)
            flr_num_txt_list.append(flr_num_txt)

            device_use = tk.Label(scrollable_body, text="    Select Device \nUse \n", font=font)
            device_use.grid(row=10, column=3)
            device_combo = ttk1.Combobox(scrollable_body, width=27)#list
            device_combo['values'] = ('Select ', 'Campus', 'Management')
            device_combo.current(0)
            device_combo.grid(row=10+i, column=3)
            device_combo_list.append(device_combo)

            l2 = tk.Label(scrollable_body, text="Enter Device \n Floor Location\n (3 Char)", font=font)
            l2.grid(row=10, column=4)

            entry_text_var_list= StringVar()
            device_floor_loc = tk.Entry(scrollable_body, width=30, textvariable=entry_text_var_list) #list
            device_floor_loc.grid(row=10 + i, column=4)
            device_floor_loc_list.append(device_floor_loc)
            entry_text_var_list.trace("w", lambda *args: charlimit3(entry_text_var_list))

            device_vendor = tk.Label(scrollable_body, text="Select the \n Device Vendor \n", font=font)
            device_vendor.grid(row=10, column=5)
            device_vendor_combo = ttk1.Combobox(scrollable_body, width=27)#list
            device_vendor_combo['values'] = ('Select ', 'Cisco', 'Cumulus', 'Lantronix', 'Arista', 'Netflow')
            device_vendor_combo.current(0)
            device_vendor_combo.grid(row=10 + i, column=5)
            device_vendor_list.append(device_vendor_combo)

            device_type = tk.Label(scrollable_body, text="    Select Device \nType\n", font=font)
            device_type.grid(row=10, column=6)
            device_type_combo = ttk1.Combobox(scrollable_body, width=27)#list
            device_type_combo['values'] = ('Select ', 'Super Spine', 'Spine', 'Core Services','Leaf',  'Terminal Server', 'Netflow', 'Layer 2')
            device_type_combo.current(0)
            device_type_combo.grid(row=10 + i, column=6)
            device_type_list.append(device_type_combo)

            device_num3 = tk.Label(scrollable_body, text="Enter Device \nPosition \n (2 Char)", font=font)
            device_num3.grid(row=10, column=7)
            device_pos = tk.Entry(scrollable_body, width=30)#list
            device_pos.grid(row=10 + i, column=7)
            device_pos_list.append(device_pos)



        #for i in device_floor_loc_list:
            #entry_text_var = i.get()
            #if len(entry_text_var)>0:
                #i.set(i.get()[:3])

            #i.configure(textvariable=entry_text_var)
            #entry_text_var.trace("w", lambda *args: charlimit3(entry_text_var))
        scrollable_body.update()



        def Config_Function():
            serialvar=0
            device_combo_var=0
            hostname_list=[]
            device_type_var_list=[]

            '''for j in device_combo_list:
                device_combo_var=j.get()
                if device_combo_var=="Campus":
                    device_combo_var="Cam"
                elif device_combo_var=="Management":
                    device_combo_var="Mgt"
                print(device_combo_var)'''
            #--------------------------------------------
            #            CREATE EXCEL File
            # --------------------------------------------
            now = datetime.now()
            runtime = 'result{}.xlsx'.format(now.strftime("%c")).replace(':', '_')
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
            ws1.set_column(1, 1, 25, normalFormat)
            ws1.set_column(2, 2, 30, normalFormat)  # Altered in SESSION2
            ws1.set_column(3, 3, 40, normalFormat)  # Altered in SESSION2
            ws1.set_column(4, 4, 15, normalFormat)  # Altered in SESSION2

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
                    device_combo_var = "Cam"
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
                ws1.write(row, 5, device_type_list[i].get())
                row += 1
            wb.close()




        Cancel_btn1 = tk.Button(root, text="BACK", command=exit)
        Cancel_btn1.place(x=0.5,y=0.9,relx=0.6,rely=0.9, anchor='n')
        Go1_btn1 = tk.Button(root, text='Export', command=Config_Function)
        Go1_btn1.place(x=0.5,y=0.9,relx=0.4,rely=0.9, anchor='n')




    Cancel_btn = tk.Button(New_NGC_page_frame, text="BACK", command=exit)
    Cancel_btn.grid(row=8 + 7 + 1, column=2)
    Go1_btn = tk.Button(New_NGC_page_frame, text='NEXT225', command=NGC_CISCO_Function)
    Go1_btn.grid(row=8 + 7 + 1, column=3)
    root.mainloop()


Cancel_btn = tk.Button(New_NGC_page_frame, text="BACK", command=exit)
Cancel_btn.grid(row=8 + 7 + 1, column=2)
Go1_btn = tk.Button(New_NGC_page_frame, text='NEXT225', command=New_NGC_Partial_Funtion)
Go1_btn.grid(row=8 + 7 + 1, column=3)
root.mainloop()

