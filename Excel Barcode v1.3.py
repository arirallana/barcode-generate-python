import xlrd
import xlwt
import math
import os
import sys
from tkinter import *
import tkinter.messagebox
from tkinter.filedialog import askopenfilename
from tkinter import filedialog
import comtypes.client
import comtypes
import collections
import win32gui
import win32api
import win32con
import urllib.request
import shutil
import ctypes
from ctypes import wintypes
from pathlib import Path
from xml.etree.ElementTree import Element, ElementTree
import xml.etree.ElementTree as etree
from xml.etree.ElementTree import SubElement, Comment, tostring
from xml.dom import minidom
try:
    import winreg
except ImportError:
    import _winreg as winreg

class MainWindow():

    def __init__(self, master):


        self.master = master
        self.row_g = tkinter.IntVar()
        self.col_g = tkinter.IntVar()
        self.a = StringVar()
        self.b = StringVar()
        self.c = StringVar()
        self.save_clickcount = 0        
       
        #Window Title
        self.master.title("Aga Khan Education Service, Pakistan - Excel Barcode Generate")
        self.master.geometry('590x580+385+25')
        self.master.resizable(False, False)
        
        #Title
        self.heading = Label(self.master, text="Excel Barcode Generate", font="timesnewroman 12 bold", fg='green', bg='white')
        self.heading.grid(row=0, column=0, columnspan=4, sticky=N+E+W+S)
        
        #Generate Frame
        self.generate_frame = LabelFrame(self.master, text="Generate", font = "timesnewroman 10 bold", labelanchor=NE)
        self.generate_frame.grid(row=2, column=0, columnspan=4, sticky=N+E+W+S)
        
        #Filename 
        Label(self.generate_frame, text="Source File Path:", font="timesnewroman 12 bold") .grid(row=3, column=0, sticky=W)
        self.filename = Entry(self.generate_frame, width=78, bg="white", textvariable=self.a)
        self.filename.grid(row=4, column=0, columnspan = 1, sticky=W)
        self.browse1 = Button(self.generate_frame, text="Browse", font="timesnewroman 12 bold",  fg="black", width=10, height=1, border = 3, relief='raised',
                             command = lambda: self.get_source_filename())
        self.browse1.grid(row= 4, column=1, sticky=W)

        #Save As
        Label(self.generate_frame, text="Save as:", font="timesnewroman 12 bold") .grid(row=5, column=0, sticky=W)
        self.save_as = Entry(self.generate_frame, width=78, bg="white", textvariable=self.b)
        self.save_as.grid(row=6, column=0, columnspan = 3, sticky=W)
        self.browse2 = Button(self.generate_frame, text="Browse", font="timesnewroman 12 bold",  fg="black", width=10, height=1, border = 3, relief='raised',
                             command = lambda: self.get_output_filename())
        self.browse2.grid(row= 6, column=1, sticky=W)

        #Sheet 
        Label(self.generate_frame, text="Sheet Name:", font="timesnewroman 12 bold") .grid(row=7, column=0, sticky=W)
        self.sheet = Entry(self.generate_frame, width=45, bg="white", textvariable=StringVar(value='Sheet1'))
        self.sheet.grid(row=8, column=0, sticky=W)

        #Start Row
        Label(self.generate_frame, text="Start Row:", font="timesnewroman 12 bold") .grid(row=9, column=0, sticky=W)
        self.start = Entry(self.generate_frame, width=20, bg="white", textvariable=IntVar())
        self.start.grid(row=10, column=0, sticky=W)

        #End Row
        Label(self.generate_frame, text="End Row:", font="timesnewroman 12 bold") .grid(row=11, column=0, sticky=W)
        self.stop = Entry(self.generate_frame, width=20, bg="white", textvariable=IntVar())
        self.stop.grid(row=12, column=0, sticky=W)

        #Barcode Column Number
        Label(self.generate_frame, text="Barcode column in source file:", font="timesnewroman 12 bold") .grid(row=13, column=0, sticky=W)
        self.col_num = Entry(self.generate_frame, width=20, bg="white", textvariable=IntVar())
        self.col_num.grid(row=14, column=0, sticky=W)

        #Call Number Checkbutton
        self.cb_displaycall_var = tkinter.IntVar()
        self.cb_displaycall = Checkbutton(self.generate_frame,text="Include CallNo.", font="timesnewroman 12 bold", variable = self.cb_displaycall_var,
                                          onvalue=True, offvalue=False, command=lambda: self.add_callno())
        self.cb_displaycall.grid(row=19, column=0, sticky=W)
        
        #Display Call Number
        self.cap_label = Label(self.generate_frame, text="CallNo. column in source file:", font="timesnewroman 12 bold")
        self.cap_num = Entry(self.generate_frame, width=20, bg="white", textvariable=IntVar())

        #Page Type
        Label(self.generate_frame, text="Select Page Type:", font="timesnewroman 12 bold") .grid(row=17, column=0, sticky=W)
        self.pagetype_v = StringVar()
        self.pagetype = OptionMenu(self.generate_frame, self.pagetype_v, 'A3', 'A4', 'A5', 'Letter', 'Legal' )
        self.pagetype.grid(row = 18, column =0, sticky=W, ipadx=20)
        self.pagetype_v.set('A4')

        #Setting Button
        self.setting_button = Button(self.master, text="SETTING", font="timesnewroman 12 bold",  fg="white", bg = "green", width=26, height=2, border = 3,
                                     relief='raised', command = lambda: self.settings_win(True))
        self.setting_button.grid(row= 27, column=0, sticky=E)

        #Generate Button
        self.bgenerate = Button(self.master, text="GENERATE", font="timesnewroman 12 bold",  fg="white", bg="green", border=3, width=26, height=2,
                                command = lambda: self.main_fieldchecker())
        self.bgenerate.grid(row= 27, column=1, sticky=E)

        #Exit Button
        self.exit = Button(self.master, text="EXIT", font="timesnewroman 12 bold",  fg="black", width=26, height=2, border = 3, relief='raised',
                           command = lambda: self.master.destroy())
        self.exit.grid(row= 28, column=0, sticky=N+S, columnspan=2)

        #Copyright
        Label(self.master, text='\u00a9 2019 Aga Khan Education Service, Pakistan All Rights Reserved', font="timesnewroman 9 bold").grid(row=30, column=0,
                                                                                                                                          columnspan=4, sticky=W)
                
        #Version
        Label(self.master, text='Version 1.3 (Last Updated on 20th August 2019)', font="timesnewroman 9 bold").grid(row=29, column=0, columnspan=4, sticky=W)

        self.filename.focus()
        self.master.withdraw()
        self.create_password_prompt()

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def install_font(self, src_path):

        user32 = ctypes.WinDLL('user32', use_last_error=True)
        gdi32 = ctypes.WinDLL('gdi32', use_last_error=True)

        FONTS_REG_PATH = r'Software\Microsoft\Windows NT\CurrentVersion\Fonts'

        HWND_BROADCAST   = 0xFFFF
        SMTO_ABORTIFHUNG = 0x0002
        WM_FONTCHANGE    = 0x001D
        GFRI_DESCRIPTION = 1
        GFRI_ISTRUETYPE  = 3

        if not hasattr(wintypes, 'LPDWORD'):
            wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)

        user32.SendMessageTimeoutW.restype = wintypes.LPVOID
        user32.SendMessageTimeoutW.argtypes = (
            wintypes.HWND,   # hWnd
            wintypes.UINT,   # Msg
            wintypes.LPVOID, # wParam
            wintypes.LPVOID, # lParam
            wintypes.UINT,   # fuFlags
            wintypes.UINT,   # uTimeout
            wintypes.LPVOID) # lpdwResult

        gdi32.AddFontResourceW.argtypes = (
            wintypes.LPCWSTR,) # lpszFilename

        # http://www.undocprint.org/winspool/getfontresourceinfo
        gdi32.GetFontResourceInfoW.argtypes = (
            wintypes.LPCWSTR, # lpszFilename
            wintypes.LPDWORD, # cbBuffer
            wintypes.LPVOID,  # lpBuffer
            wintypes.DWORD)   # dwQueryType

        # copy the font to the Windows Fonts folder
        dst_path = os.path.join(os.environ['SystemRoot'], 'Fonts',
                                os.path.basename(src_path))
        dst_path = dst_path.replace("\\", "/")
        shutil.copy(src_path, dst_path)
        # load the font in the current session
        if not gdi32.AddFontResourceW(dst_path):
            os.remove(dst_path)
            raise WindowsError('AddFontResource failed to load "%s"' % src_path)
        # notify running programs
        user32.SendMessageTimeoutW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0,
                                   SMTO_ABORTIFHUNG, 1000, None)
        # store the fontname/filename in the registry
        filename = os.path.basename(dst_path)
        fontname = os.path.splitext(filename)[0]
        # try to get the font's real name
        cb = wintypes.DWORD()
        if gdi32.GetFontResourceInfoW(filename, ctypes.byref(cb), None,
                                      GFRI_DESCRIPTION):
            buf = (ctypes.c_wchar * cb.value)()
            if gdi32.GetFontResourceInfoW(filename, ctypes.byref(cb), buf,
                                          GFRI_DESCRIPTION):
                fontname = buf.value
        is_truetype = wintypes.BOOL()
        cb.value = ctypes.sizeof(is_truetype)
        gdi32.GetFontResourceInfoW(filename, ctypes.byref(cb),
            ctypes.byref(is_truetype), GFRI_ISTRUETYPE)
        if is_truetype:
            fontname += ' (TrueType)'
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, FONTS_REG_PATH, 0,
                            winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, fontname, 0, winreg.REG_SZ, filename)


    def font_check(self, fontname=str):
        fontslist = []
        res = self.enum_fonts()
        
        for r in res:
            fontslist.append(r[0].lfFaceName)

        print(fontslist)
        print(fontname)

        if fontname not in fontslist:
            self.install_font(self.resource_path('code128.ttf'))

    def enum_fonts(self, typeface=None):
        hwnd = win32gui.GetDesktopWindow()
        dc = win32gui.GetWindowDC(hwnd)

        res = []
        def callback(*args):
            res.append(args)
            return 1
        win32gui.EnumFontFamilies(dc, typeface, callback)

        win32gui.ReleaseDC(hwnd, dc)
        return res

    def add_callno(self):
        position = "385+25"
        if hasattr(self, "root2") and self.root2.winfo_exists():
            if self.root2.winfo_ismapped():
                position = "10+25"   
        if bool(self.cb_displaycall_var.get()):
            self.master.geometry('590x625+'+position)
            self.cap_label.grid(row=20, column=0, sticky=W)
            self.cap_num.grid(row=21, column=0, sticky=W)
        else:
            self.master.geometry('590x580+'+position)
            self.cap_label.grid_forget()
            self.cap_num.grid_forget()

    def create_password_prompt(self):
        self.root3 = Toplevel(self.master)        
        self.root3.geometry('280x120+535+250')
        self.root3.resizable(False, False)

        #Window Title
        self.root3.title("Password Prompt")

        #Password Prompt Frame
        self.password_frame = LabelFrame(self.root3, text="Password Prompt", font = "timesnewroman 10 bold", labelanchor=NE)
        self.password_frame.grid(row=2, column=0, columnspan=2, sticky=N+E+W+S)

        #Enter Password
        Label(self.password_frame, text="Enter Password:", font="timesnewroman 12 bold") .grid(row=3, column=0, sticky=W, columnspan=2)
        self.password = Entry(self.password_frame, width=45, bg="white", textvariable=StringVar(), show="*")
        self.password.grid(row=4, column=0, sticky=W, columnspan=2)

        #Version
        Label(self.root3, text='Version 1.3 (Last Updated on 20th August 2019)', font="timesnewroman 9 bold").grid(row=29, column=0, columnspan=4, sticky=W)

        #OK Button
        self.ok = Button(self.root3, text="OK", font="timesnewroman 12 bold",  fg="black", border=3, width=10, height=1,
                                command = lambda: self.check_password())
        self.ok.grid(row= 7, column=1, sticky=W)

        #Close Button
        self.close = Button(self.root3, text="Close", font="timesnewroman 12 bold",  fg="black", border=3, width=10, height=1,
                                command = lambda: self.root3.destroy())
        self.close.grid(row= 7, column=0, sticky=E)
            
        self.password.focus()
        
    def settings_win(self, visible=bool):
        
        if hasattr(self, "root2") and self.root2.winfo_exists():
            self.settings_not_changed()
            if visible:
                if bool(self.cb_displaycall_var.get()):
                    self.master.geometry('590x625+10+25')
                    self.master.update()
                else:
                    self.master.geometry('590x580+10+25')
                    self.master.update()
                self.root2.deiconify()
                self.col_lim.focus()
            else:
                self.root2.withdraw()
        else:
            self.root2 = Toplevel(self.master)
            
            if bool(self.cb_displaycall_var.get()):
                self.row_g.set(1)
            else:
                self.row_g.set(0)
            
            if visible:
                self.root2.deiconify()
                if bool(self.cb_displaycall_var.get()):
                    self.master.geometry('590x625+10+25')
                    self.master.update()
                else:
                    self.master.geometry('590x580+10+25')
                    self.master.update()
            else:
                self.root2.withdraw()
                
            self.root2.geometry('350x420+620+25')
            self.root2.resizable(False, False)

            #Window Title
            self.root2.title("Excel Barcode Generate - Settings")

            #Settings Frame
            self.settings_frame = LabelFrame(self.root2, text="Settings", font = "timesnewroman 11 bold", labelanchor=NE)
            self.settings_frame.grid(row=0, column=0, columnspan=2, sticky=N+E+W+S)

            #Title
            self.heading = Label(self.settings_frame, text="Excel Barcode Generate", font="timesnewroman 11 bold", fg='green', bg='white')
            self.heading.grid(row=1, column=0, columnspan=2, sticky=N+E+W+S)
            self.subheading = Label(self.settings_frame, text="Settings", font="timesnewroman 11 bold", fg='green', bg='white')
            self.subheading.grid(row=2, column=0, columnspan=2, sticky=N+E+W+S)            

            #Columns per Page
            Label(self.settings_frame, text="Columns per Page:", font="timesnewroman 11 bold") .grid(row=3, column=0, sticky=W)
            self.col_lim = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.col_lim.grid(row=4, column=0, sticky=W)

            #Row Height
            Label(self.settings_frame, text="Row Height:", font="timesnewroman 11 bold") .grid(row=5, column=0, sticky=W)
            self.row_height = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.row_height.grid(row=6, column=0, sticky=W)

            #Column Width
            Label(self.settings_frame, text="Column Width:", font="timesnewroman 11 bold") .grid(row=7, column=0, sticky=W)
            self.col_width = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.col_width.grid(row=8, column=0, sticky=W)

            #Barcode Font
            Label(self.settings_frame, text="Barcode Font Size:", font="timesnewroman 11 bold") .grid(row=9, column=0, sticky=W)
            self.barcode_font = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.barcode_font.grid(row=10, column=0, sticky=W)

            #Text Font
            Label(self.settings_frame, text="Text Font Size:", font="timesnewroman 11 bold") .grid(row=11, column=0, sticky=W)
            self.text_font = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.text_font.grid(row=12, column=0, sticky=W)

            #Row Gap
            self.row_gap = Checkbutton(self.settings_frame, text = "Add Row Gap",variable = self.row_g, onvalue = 1, offvalue = 0, 
                                       font="timesnewroman 11 bold", command = lambda: self.add_row_gap())
            self.row_gap.grid(row=13, column=0, sticky=W)
            self.row_gap_label = Label(self.settings_frame, text="Row Gap Height:", font="timesnewroman 11 bold") 
            self.row_gap_height = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            
            #Column Gap
            self.col_gap = Checkbutton(self.settings_frame, text = "Add Column Gap", variable = self.col_g, onvalue = 1, offvalue = 0,
                                       font="timesnewroman 11 bold", command = lambda: self.add_col_gap())
            self.col_gap.grid(row=16, column=0, sticky=W)
            self.col_gap_label = Label(self.settings_frame, text="Column Gap Width:", font="timesnewroman 11 bold") 
            self.col_gap_width = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())

            #Page Margins
            Label(self.settings_frame, text="Set Page Margins (inches):", font="timesnewroman 11 bold") .grid(row=4, column=1, sticky=W)
            
            Label(self.settings_frame, text="Left:", font="timesnewroman 11 bold") .grid(row=5, column=1, sticky=W)
            self.l_m = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.l_m.grid(row=6, column=1, sticky=W)

            Label(self.settings_frame, text="Right:", font="timesnewroman 11 bold") .grid(row=7, column=1, sticky=W)
            self.r_m = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.r_m.grid(row=8, column=1, sticky=W)

            Label(self.settings_frame, text="Top:", font="timesnewroman 11 bold") .grid(row=9, column=1, sticky=W)
            self.t_m = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.t_m.grid(row=10, column=1, sticky=W)

            Label(self.settings_frame, text="Bottom:", font="timesnewroman 11 bold") .grid(row=11, column=1, sticky=W)
            self.b_m = Entry(self.settings_frame, width=20, bg="white", textvariable=IntVar())
            self.b_m.grid(row=12, column=1, sticky=W)

            #Save Button
            self.save = Button(self.settings_frame, text="SAVE", font="timesnewroman 11 bold",  fg="white", bg="green", border=3, width=10, height=1,
                                    command = lambda: self.settings_fieldchecker())
            self.save.grid(row= 28, column=1, sticky=E)

            #Exit Button
            self.exit2 = Button(self.settings_frame, text="EXIT", font="timesnewroman 11 bold",  fg="black", border=3, width=10, height=1,
                                    command = lambda: self.root2_exit())
            self.exit2.grid(row= 28, column=1, sticky=W)

            #Reset Button
            self.reset = Button(self.settings_frame, text="RESET", font="timesnewroman 11 bold",  fg="white", bg="green", border=3, width=10, height=1,
                                    command = lambda: self.reset_settings())
            self.reset.grid(row= 27, column=0, sticky=W)

            #Clear Saves Button
            self.default = Button(self.settings_frame, text="CLEAR SAVES", font="timesnewroman 11 bold",  fg="white", bg="green", border=3, width=15, height=1,
                                    command = lambda: self.clear_saves())
            self.default.grid(row= 28, column=0, sticky=W)

            self.col_lim.focus()
            self.root2.protocol("WM_DELETE_WINDOW", lambda: self.on_closing())

            self.set_defaults(self.pagetype_v.get())
            
        self.read_saved_settings()
        print(self.saved_in_xml)
        if len(self.saved_in_xml) != 0:
            self.set_saved(self.saved_in_xml)

        self.add_row_gap()
        self.add_col_gap()
        self.settings_win_sizer()

        self.saved_values = self.get_current_values()

    def clear_saves(self):
        if bool(self.save_clickcount) or os.path.exists('saved_settings.xml'):
            tkinter.messagebox.showinfo("Notification", "Previous save was cleared. All settings restored to default. ")
            self.save_clickcount=0
            self.root2.destroy()
            self.master.update()
            self.add_callno()
            try:
                os.remove('saved_settings.xml')
            except FileNotFoundError:
                pass
        else:
            tkinter.messagebox.showinfo("Notification", "No previous save. ")
        
    def reset_settings(self):
        if bool(self.save_clickcount):
            self.set_saved_settings()
        else:
            self.set_defaults(self.pagetype_v.get())

    def add_row_gap(self):
        if bool(self.row_g.get()):
            self.row_gap_label.grid(row=14, column=0, sticky=W)
            self.row_gap_height.grid(row=15, column=0, sticky=W)
            self.root2.geometry('350x460+620+25')
        else:
            self.row_gap_label.grid_forget()
            self.row_gap_height.grid_forget()
            self.root2.geometry('350x420+620+25')

        self.settings_win_sizer()

    def add_col_gap(self):
        if bool(self.col_g.get()):
            self.col_gap_label.grid(row=17, column=0, sticky=W)
            self.col_gap_width.grid(row=18, column=0, sticky=W)
        else:
            self.col_gap_label.grid_forget()
            self.col_gap_width.grid_forget()

        self.settings_win_sizer()

    def settings_win_sizer(self):
        if bool(self.row_g.get()) and bool(self.col_g.get()):
            self.root2.geometry('350x500+620+25')
        elif bool(self.row_g.get()) or bool(self.col_g.get()):
            self.root2.geometry('350x460+620+25')
        else:
            self.root2.geometry('350x420+620+25')

    def set_saved(self, saved_in_xml):
        self.col_lim.delete(0, END)
        self.col_lim.insert(0, saved_in_xml[0])
        self.row_height.delete(0, END)
        self.row_height.insert(0, saved_in_xml[1])
        self.col_width.delete(0, END)
        self.col_width.insert(0, saved_in_xml[2])
        self.barcode_font.delete(0, END)
        self.barcode_font.insert(0, saved_in_xml[3])
        self.text_font.delete(0, END)
        self.text_font.insert(0, saved_in_xml[4])
        self.row_g.set(saved_in_xml[5])
        self.col_g.set(saved_in_xml[6])
        self.l_m.delete(0, END)
        self.l_m.insert(0, saved_in_xml[7])
        self.r_m.delete(0, END)
        self.r_m.insert(0, saved_in_xml[8])
        self.b_m.delete(0, END)
        self.b_m.insert(0, saved_in_xml[9])
        self.t_m.delete(0, END)
        self.t_m.insert(0, saved_in_xml[10])
        self.row_gap_height.delete(0, END)
        self.row_gap_height.insert(0, saved_in_xml[11])
        self.col_gap_width.delete(0, END)
        self.col_gap_width.insert(0, saved_in_xml[12])
            
    def set_defaults(self, pagetype=str):
        if bool(self.cb_displaycall_var.get()):
            self.row_g.set(1)
        else:
            self.row_g.set(0)
        self.col_g.set(0)
        self.add_row_gap()
        self.add_col_gap()
        self.settings_win_sizer()
        
        if pagetype == 'A4':
            self.col_lim.delete(0, END)
            self.col_lim.insert(0, 4)
            self.row_height.delete(0, END)
            self.row_height.insert(0, 3)
            self.col_width.delete(0, END)
            self.col_width.insert(0, 35)
            self.barcode_font.delete(0, END)
            self.barcode_font.insert(0, 27)
            self.text_font.delete(0, END)
            self.text_font.insert(0, 12)
            self.l_m.delete(0, END)
            self.l_m.insert(0, 0)
            self.r_m.delete(0, END)
            self.r_m.insert(0, 0)
            self.b_m.delete(0, END)
            self.b_m.insert(0, 0)
            self.t_m.delete(0, END)
            self.t_m.insert(0, 0.5)
            self.row_gap_height.delete(0, END)
            self.row_gap_height.insert(0, 1)
            self.col_gap_width.delete(0, END)
            self.col_gap_width.insert(0, 5)

        elif pagetype == 'A3':
            self.col_lim.delete(0, END)
            self.col_lim.insert(0, 4)
            self.row_height.delete(0, END)
            self.row_height.insert(0, 4)
            self.col_width.delete(0, END)
            self.col_width.insert(0, 35)
            self.barcode_font.delete(0, END)
            self.barcode_font.insert(0, 30)
            self.text_font.delete(0, END)
            self.text_font.insert(0, 13)
            self.l_m.delete(0, END)
            self.l_m.insert(0, 0)
            self.r_m.delete(0, END)
            self.r_m.insert(0, 0)
            self.t_m.delete(0, END)
            self.t_m.insert(0, 0.4)
            self.b_m.delete(0, END)
            self.b_m.insert(0, 0.5)
            self.row_gap_height.delete(0, END)
            self.row_gap_height.insert(0, 1)
            self.col_gap_width.delete(0, END)
            self.col_gap_width.insert(0, 5)

        elif pagetype == 'A5':
            self.col_lim.delete(0, END)
            self.col_lim.insert(0, 4)
            self.row_height.delete(0, END)
            self.row_height.insert(0, 3)
            self.col_width.delete(0, END)
            self.col_width.insert(0, 35)
            self.barcode_font.delete(0, END)
            self.barcode_font.insert(0, 28)
            self.text_font.delete(0, END)
            self.text_font.insert(0, 13)
            self.l_m.delete(0, END)
            self.l_m.insert(0, 0)
            self.r_m.delete(0, END)
            self.r_m.insert(0, 0)
            self.t_m.delete(0, END)
            self.t_m.insert(0, 0.6)
            self.b_m.delete(0, END)
            self.b_m.insert(0, 0.6)
            self.row_gap_height.delete(0, END)
            self.row_gap_height.insert(0, 1)
            self.col_gap_width.delete(0, END)
            self.col_gap_width.insert(0, 5)

        elif pagetype == 'Letter':
            self.col_lim.delete(0, END)
            self.col_lim.insert(0, 4)
            self.row_height.delete(0, END)
            self.row_height.insert(0, 3)
            self.col_width.delete(0, END)
            self.col_width.insert(0, 35)
            self.barcode_font.delete(0, END)
            self.barcode_font.insert(0, 27)
            self.text_font.delete(0, END)
            self.text_font.insert(0, 12)
            self.l_m.delete(0, END)
            self.l_m.insert(0, 0)
            self.r_m.delete(0, END)
            self.r_m.insert(0, 0)
            self.b_m.delete(0, END)
            self.b_m.insert(0, 0)
            self.t_m.delete(0, END)
            self.t_m.insert(0, 0.5)
            self.row_gap_height.delete(0, END)
            self.row_gap_height.insert(0, 1)
            self.col_gap_width.delete(0, END)
            self.col_gap_width.insert(0, 5)

        elif pagetype == 'Legal':
            self.col_lim.delete(0, END)
            self.col_lim.insert(0, 4)
            self.row_height.delete(0, END)
            self.row_height.insert(0, 3)
            self.col_width.delete(0, END)
            self.col_width.insert(0, 35)
            self.barcode_font.delete(0, END)
            self.barcode_font.insert(0, 27)
            self.text_font.delete(0, END)
            self.text_font.insert(0, 12)
            self.l_m.delete(0, END)
            self.l_m.insert(0, 0)
            self.r_m.delete(0, END)
            self.r_m.insert(0, 0)
            self.t_m.delete(0, END)
            self.t_m.insert(0, 0.5)
            self.b_m.delete(0, END)
            self.b_m.insert(0, 3.2)
            self.row_gap_height.delete(0, END)
            self.row_gap_height.insert(0, 1)
            self.col_gap_width.delete(0, END)
            self.col_gap_width.insert(0, 5)

        if not bool(self.cb_displaycall_var.get()):
            if pagetype == 'A4':
                self.b_m.delete(0, END)
                self.b_m.insert(0, 1.5)
                self.t_m.delete(0, END)
                self.t_m.insert(0, 0.5)
            if pagetype == 'A3':
                self.b_m.delete(0, END)
                self.b_m.insert(0, 1.2)
                self.t_m.delete(0, END)
                self.t_m.insert(0, 0.5)
            if pagetype == 'A5':
                self.b_m.delete(0, END)
                self.b_m.insert(0, 1.8)
                self.t_m.delete(0, END)
                self.t_m.insert(0, 0.5)
            if pagetype == 'Letter':
                self.b_m.delete(0, END)
                self.b_m.insert(0, 1.5)
                self.t_m.delete(0, END)
                self.t_m.insert(0, 0.5)
            if pagetype == 'Legal':
                self.b_m.delete(0, END)
                self.b_m.insert(0, 4.5)
                
    def get_current_values(self):
        return [self.col_lim.get(), self.row_height.get(), self.col_width.get(), self.barcode_font.get(),
                self.text_font.get(), self.row_g.get(), self.col_g.get(), self.l_m.get(), self.r_m.get(),
                self.t_m.get(), self.b_m.get(), self.row_gap_height.get(), self.col_gap_width.get()]

    def on_closing(self):
        if bool(self.cb_displaycall_var.get()):
            self.master.geometry('590x625+385+25')
            self.master.update()
        else:
            self.master.geometry('590x580+385+25')
            self.master.update()
        self.root2.withdraw()

    def set_saved_settings(self):
        self.col_lim.delete(0, END)
        self.col_lim.insert(0, self.saved_values[0])
        
        self.row_height.delete(0, END)
        self.row_height.insert(0, self.saved_values[1])

        self.col_width.delete(0, END)
        self.col_width.insert(0, self.saved_values[2])

        self.barcode_font.delete(0, END)
        self.barcode_font.insert(0, self.saved_values[3])

        self.text_font.delete(0, END)
        self.text_font.insert(0, self.saved_values[4])

        self.row_g.set(int(self.saved_values[5]))
        self.col_g.set(int(self.saved_values[6]))
        
        self.l_m.delete(0, END)
        self.l_m.insert(0, self.saved_values[7])

        self.r_m.delete(0, END)
        self.r_m.insert(0, self.saved_values[8])

        self.t_m.delete(0, END)
        self.t_m.insert(0, self.saved_values[9])

        self.b_m.delete(0, END)
        self.b_m.insert(0, self.saved_values[10])

        self.row_gap_height.delete(0, END)
        self.row_gap_height.insert(0, self.saved_values[11])

        self.col_gap_width.delete(0, END)
        self.col_gap_width.insert(0, self.saved_values[12])
       
        self.add_row_gap()
        self.add_col_gap()
        self.settings_win_sizer()

    def edit_check(self):
        self.current_values =  self.get_current_values()
        if collections.Counter(self.current_values) != collections.Counter(self.saved_values):
            return True
        else:
            return False
        
    def root2_exit(self):
        if self.edit_check():
            question = tkinter.messagebox.askquestion ('Exit Settings','Are you sure you want to exit without saving?',icon = 'warning')
            if question == 'yes':
               self.set_saved_settings()
               self.on_closing()
            else:
                self.col_lim.focus()
        else:
            self.root2.withdraw()
            self.on_closing()
            
    def check_password(self):
        if self.password.get() == 'akesp':
            self.root3.destroy()
            self.master.deiconify()
            self.font_check('Code 128')
        else:
            tkinter.messagebox.showinfo("Notification", "Incorrect password. ")
            self.root3.destroy()
            self.master.destroy()

    def settings_not_changed(self):
        if not bool(self.save_clickcount):
                self.set_defaults(self.pagetype_v.get())

    def execute_bgenerate(self):
        self.task_kill()
        self.settings_not_changed()
        Code128().barcode_generate(self.filename.get(), self.start.get(), self.stop.get(),
                                                                             self.sheet.get(), self.col_num.get(), self.cap_num.get(),
                                                                             self.col_lim.get(), self.row_height.get(), self.col_width.get(),
                                                                             self.barcode_font.get(), self.text_font.get(),
                                                                             self.row_g.get(), self.col_g.get(), self.l_m.get(),
                                                                             self.r_m.get(), self.t_m.get(), self.b_m.get(),
                                                                             self.save_as.get(), self.cb_displaycall_var.get(), self.pagetype_v.get(),
                                                                             self.row_gap_height.get(), self.col_gap_width.get())
        print(self.get_current_values())
        self.root2.withdraw()
        
            
    def check_settings(self):
        try:
            self.execute_bgenerate()
        except AttributeError:
            self.settings_win(False)
            self.execute_bgenerate()           
        
    def special_char(self, s):
        chars = set('[*?/\:]')
        if any((c in chars) for c in s):
            return True
        else:
            return False

    def main_filechecker(self):
        try:
            source_workbook = xlrd.open_workbook(self.filename.get())
            try:
                source_worksheet = source_workbook.sheet_by_name(self.sheet.get())
                print(source_worksheet.ncols)
                print(source_worksheet.nrows)
                if int(self.col_num.get()) >source_worksheet.ncols:
                    tkinter.messagebox.showinfo("Error", "Barcode column not found. ")
                elif int(self.cap_num.get()) >source_worksheet.ncols:
                    tkinter.messagebox.showinfo("Error", "CallNo. column not found. ")
                elif int(self.start.get())>source_worksheet.nrows:
                    tkinter.messagebox.showinfo("Error", "Start row not found. ")
                elif int(self.stop.get())>source_worksheet.nrows:
                    tkinter.messagebox.showinfo("Error", "End row not found. ")
                else:
                    self.check_settings()
            except xlrd.XLRDError:
                tkinter.messagebox.showinfo("Error", "Sheet not found. ")
        except IOError:
            tkinter.messagebox.showinfo("Error", "File or Directory not found. ")
        
    def main_fieldchecker(self):
        try:
            if hasattr(self, "root2") and self.root2.winfo_ismapped():
                tkinter.messagebox.showinfo("Error", "Please SAVE or EXIT Settings then click Generate. ")
            elif not os.path.exists(os.path.dirname(self.filename.get())):
                tkinter.messagebox.showinfo("Error", "Please enter valid file path. ")
            elif not os.path.exists(os.path.dirname(self.save_as.get())):
                tkinter.messagebox.showinfo("Error", "Please enter valid save path. ")
            elif (self.sheet.get() == '') or (len(self.sheet.get())>31) or (self.special_char(self.sheet.get())):
                tkinter.messagebox.showinfo("Error", "Please enter valid sheet name. ")
            elif int(self.start.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid start row. ")
            elif int(self.stop.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid end row. ")
            elif int(self.col_num.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid barcode column number. ")
            elif bool(self.cb_displaycall_var.get()) and (int(self.cap_num.get()) < 1):
                tkinter.messagebox.showinfo("Error", "Please enter valid CallNo column number. ")
            else:
                self.main_filechecker()
        except ValueError:
            tkinter.messagebox.showinfo("Error", "Numeric values must be positive integers. ")
        
    def settings_fieldchecker(self):
        try:
            if int(self.col_lim.get()) < 1 :
                tkinter.messagebox.showinfo("Error", "Please enter valid number of columns. ")
            elif int(self.row_height.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid row height. ")
            elif int(self.col_width.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid column width. ")
            elif int(self.barcode_font.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid barcode font size. ")
            elif int(self.text_font.get()) < 1:
                tkinter.messagebox.showinfo("Error", "Please enter valid text font size. ")
            elif bool(self.l_m.get()) < 0 or not self.l_m.get():
                tkinter.messagebox.showinfo("Error", "Please enter valid left margin size. ")
            elif bool(self.r_m.get()) < 0 or not self.r_m.get():
                tkinter.messagebox.showinfo("Error", "Please enter valid right margin size. ")
            elif bool(self.t_m.get()) < 0 or not self.t_m.get():
                tkinter.messagebox.showinfo("Error", "Please enter valid top margin size. ")
            elif bool(self.b_m.get()) < 0 or not self.b_m.get():
                tkinter.messagebox.showinfo("Error", "Please enter valid bottom margin size. ")
            elif int(self.col_gap_width.get())<1:
                tkinter.messagebox.showinfo("Error", "Please enter valid column gap width. ")
            elif int(self.row_gap_height.get())<1:
                tkinter.messagebox.showinfo("Error", "Please enter valid row gap height. ")
            else:
                if bool(self.cb_displaycall_var.get()):
                    self.master.geometry('590x625+385+25')
                    self.master.update()
                else:
                    self.master.geometry('590x580+385+25')
                    self.master.update()
                if self.edit_check():
                    self.save_clickcount+=1
                self.root2.withdraw()
                print(self.save_clickcount)
                self.save_as_xml()
        except ValueError:
            tkinter.messagebox.showinfo("Error", "Numeric values must be positive integers. ")
            self.reset_settings()
            
    def read_saved_settings(self):
        self.saved_in_xml = []
        try:
            file = etree.parse(r'saved_settings.xml')
            root = file.getroot()
            for i in range(13):
                value = root[i].text
                self.saved_in_xml.append(value)
        except FileNotFoundError:
            pass
        return self.saved_in_xml
            
    def save_as_xml(self):
        self.saved_values = self.get_current_values()
        root = Element('saved_settings')
        tree = ElementTree(root)
        
        col_lim = Element('col_lim')
        root.append(col_lim)
        col_lim.text = self.saved_values[0]

        row_height = Element('row_height')
        root.append(row_height)
        row_height.text = self.saved_values[1]
        
        col_width = Element('col_width')
        root.append(col_width)
        col_width.text = self.saved_values[2]
        
        barcode_font = Element('barcode_font')
        root.append(barcode_font)
        barcode_font.text = self.saved_values[3]
        
        text_font = Element('text_font')
        root.append(text_font)
        text_font.text = self.saved_values[4]
        
        row_g = Element('row_g')
        root.append(row_g)
        row_g.text = str(self.saved_values[5])
        
        col_g = Element('col_g')
        root.append(col_g)
        col_g.text = str(self.saved_values[6])
        
        l_m = Element('l_m')
        root.append(l_m)
        l_m.text = self.saved_values[7]
        
        r_m = Element('r_m')
        root.append(r_m)
        r_m.text = self.saved_values[8]
        
        t_m = Element('t_m')
        root.append(t_m)
        t_m.text = self.saved_values[9]
        
        b_m = Element('b_m')
        root.append(b_m)
        b_m.text = self.saved_values[10]
        
        row_gap_height = Element('row_gap_height')
        root.append(row_gap_height)
        row_gap_height.text = self.saved_values[11]
        
        col_gap_width = Element('col_gap_width')
        root.append(col_gap_width)
        col_gap_width.text = self.saved_values[12]

        file = open(os.path.join(os.path.curdir,'saved_settings.xml'), 'w')
        tree = tree.getroot()
        file.write(etree.tostring(tree, encoding="unicode"))
        file.close()

        self.read_saved_settings()

    def task_kill(self):
        os.system('TASKKILL /F /IM excel.exe')
        os.system('TASKKILL /F /IM acrobat.exe')                  

    def get_source_filename(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
        self.a.set(askopenfilename(initialdir = desktop, title = "Select file", filetypes=(("All files", "*.*"),("Excel files", "*.xlsx"))))

    def get_output_filename(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
        self.b.set(filedialog.asksaveasfilename(initialdir = desktop, title = "Save as",filetypes = ( ("Excel files","*.xls"), ("All files","*.*"))))

    def get_install_filename(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
        self.c.set(askopenfilename(initialdir = desktop, title = "Select file", filetypes=(("All files", "*.*"),("True Type Font files", "*.ttf"))))
        
class Code128:

    def list_join(self, seq):
                    
        return [x for sub in seq for x in sub]

    def encode128(self, s):
        
        code128B_mapping = dict((chr(c), [98, c+64] if c < 32 else [c-32]) for c in range(128))
        code128C_mapping = dict([(u'%02d' % i, [i]) for i in range(100)] + [(u'%d' % i, [100, 16+i]) for i in range(10)])
        code128_chars = u''.join(chr(c) for c in [212] + list(range(33,126+1)) + list(range(200,211+1)))
        
        if s.isdigit() and len(s) >= 2:
            codes = [105] + list_join(code128C_mapping[s[i:i+2]] for i in range(0, len(s), 2))
        else:
            codes = [104] + self.list_join(code128B_mapping[c] for c in s)
        check_digit = (codes[0] + sum(i * x for i,x in enumerate(codes))) % 103
        codes.append(check_digit)
        codes.append(106) 
        return u''.join(code128_chars[x] for x in codes)

    def barcode_generate(self, filename, start, stop, sheet, col_num, cap_num, col_lim, row_height, col_width, barcode_font,
                         text_font, row_gap, col_gap, l_m, r_m, t_m, b_m, save_as, has_callno, paper_size, row_gap_height, col_gap_width):

        start = int(start)
        stop = int(stop)
        col_lim = int(col_lim)        

        source_workbook = xlrd.open_workbook(filename)
        source_worksheet = source_workbook.sheet_by_name(sheet)

        wb = xlwt.Workbook(save_as + '.xls')
        ws = wb.add_sheet('Barcodes '+str(start)+'-'+str(stop))
        ws._cell_overwrite_ok=True

        #Barcode Style
        barcodefont = xlwt.Font()
        barcodefont.name = 'Code 128'
        barcodefont.height = int(barcode_font)*20
        style1 = xlwt.XFStyle()
        style1.font = barcodefont

        alignment = xlwt.Alignment()
        alignment.horz = 2 
        alignment.vert = 2
        style1.alignment = alignment

        #Value Style
        valfont = xlwt.Font()
        valfont.name = 'Arial'
        valfont.height = int(text_font)*20 
        style2 = xlwt.XFStyle()
        style2.font = valfont

        alignment = xlwt.Alignment()
        alignment.horz = 2 
        alignment.vert = 2
        style2.alignment = alignment

        #Caption Style
        capfont = xlwt.Font()
        capfont.name = 'Arial'
        capfont.height = int(text_font)*20
        style3 = xlwt.XFStyle()
        style3.font = capfont

        alignment = xlwt.Alignment()
        alignment.horz = 2 
        alignment.vert = 2
        style3.alignment = alignment
        
            
        #Writer
        if bool(row_gap):
            step_r = 4
        else:
            step_r = 3

        if bool(col_gap):
            step_c = 2
            col_lim = col_lim*2
            end = (math.ceil((stop+1-start)/col_lim)*8)-1
        else:
            step_c = 1
            end = (math.ceil((stop+1-start)/col_lim)*4)-1

        
        count = 0             
        for row in range(0,end,step_r):      
            for col in range(0,int(col_lim),step_c):
                if count<=(stop-start):
                    value =  source_worksheet.cell(count+start-1, int(col_num)-1).value
                    value_encoded = self.encode128(str(value))
                    print(value_encoded)
                    ws.write(row+1, col, value_encoded, style1)
                    ws.write(row+2, col, value, style2)
                    caption = source_worksheet.cell(count+start-1, int(cap_num)-1).value
                    print(caption)
                    if int(cap_num)>0 and bool(has_callno):
                        ws.write(row, col, caption, style3)
                count += 1

        #Width/Height Adjustment
        if int(cap_num)>0 and bool(has_callno):
            row_start = 0
            row_gap_start = 3
            row_g_step = 4
        else:
            row_start = 0
            row_gap_start = 2
            row_g_step = 3
            
        for i in range(row_start, stop+1-start, row_g_step):
            ws.row(i).height_mismatch = True
            ws.row(i).height = 256*int(row_height)#256 is width of 0
        if bool(row_gap):
            for i in range(row_gap_start, stop+1-start, row_g_step):
                ws.row(i).height_mismatch = True
                ws.row(i).height = 256*int(row_gap_height)
        for i in range(0,col_lim):
            ws.col(i).width_mismatch = True
            ws.col(i).width = 256*int(col_width)
        if bool(col_gap):
            for i in range(1, col_lim, 2):
                ws.col(i).width_mismatch = True
                ws.col(i).width = 256*int(col_gap_width)

        print(l_m)
        print(r_m)
        print(t_m)
        print(b_m)

        papersizes = {'A3':8, 'A4':9, 'A5':11, 'Letter':1, 'Legal':5} #(see https://open-hea.readthedocs.io/en/latest/xlwt.html)

        #Page Setup
        ws.paper_size_code = papersizes[paper_size]      
        ws.left_margin = float(l_m)         # in inch
        ws.right_margin = float(r_m)
        ws.top_margin = float(t_m)
        ws.bottom_margin = float(b_m)
        ws.portrait = True
        ws.header_str = bytes("", 'utf-8')
        ws.footer_str = bytes("", 'utf-8')
        ws.header_margin = 0 
        ws.footer_margin = 0
        ws.portrait = 1
        ws.fit_num_pages = 1
        ws.fit_height_to_pages = 0
        ws.fit_width_to_pages = 1

        print(save_as)
        path = os.path.dirname(save_as)
        print(path)

        os.chdir(path)
        
        if save_as.endswith('.xls'):
            directory, file = os.path.split(save_as)
            print(file)
            wb.save(save_as)
        else:
            directory, file = os.path.split(save_as)
            file = file+'.xls'
            print(file)
            wb.save(save_as+'.xls')

        app = comtypes.client.CreateObject('Excel.Application')
        app.Visible = False

        self.without_extension = str(file.split('.')[0])

        infile = os.path.join(os.path.abspath(os.path.curdir), str(file))
        outfile = os.path.join(os.path.abspath(os.path.curdir), self.without_extension+'.pdf')

        doc = app.Workbooks.Open(infile)
        doc.ExportAsFixedFormat(0, outfile, 1, 0)
        doc.Close()

        app.Quit()

        tkinter.messagebox.showinfo("Notification", "Job Completed. Please check output folder. ")

def main():
    root = Tk()
    myMainWindow = MainWindow(root)
    root.mainloop()

if __name__ == '__main__':
    main()

    
  
    
