# coding=utf-8
__author__ = 'yaojian'

import Tkinter, Tkconstants, tkFileDialog, tkMessageBox
import os

DMCTools = "DMCTools9.xlam"

USR_XLSTART = "C:\\Users\\yao20\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART"

Help = """
###############################################################################
#                      GDO Comparator Tool                                    #
###############################################################################
Instruction to Users:
GDO Comparator Tool is used to compare a new In-Life Report outputs with those
from previous run for identifying any new or updated data on the In-Life report
outputs during Data Integrity Review. This tool could also help to consolidate comments added by CDMs in different files into one file.   Please contact your local liaisons or
trainer if you have any issue or question about this tool
###############################################################################
The following are the introduction of each available button in the tool

Old Excel File
This button is used to select the Old Excel file for comparison.
Note: This tool only supports file in "Excel 97-2003" type. If the file is in
 other Excel type, please convert it as "Excel 97-2003 workbook" type either
using the "Consolidation" button in Excel or do it manually.

New Excel File
This button is used to select the New Excel file for comparison.
Note: This tool only supports file in "Excel 97-2003" type. If the file is in
other Excel type, please convert it as "Excel 97-2003 workbook" type
either using the "Consolidation" button in Excel or do it manually.

Output Excel File
This button is used to select a designated location to save the results after
comparison and to name the comparison result file.
Note: This must be selected before Run Comparing.

Run Comparing
This button is used to initiate the comparison program and save the comparison
result to the previously defined location. Error message and error list will
show if any report is missing in old file or new file. New outputs will be
highlighted in green and new report will be highlighted in yellow.

Cumulative Compare
This button is used to execute cumulative compare for external data after Old Excel File, New Excel File & Output Excel File are selected. The tool will highlight new records in new file, carry over comment from old excel file and also carry over records that are disappeared from old file.

Main Consolidate File
The button is used to select the main files based on which the comments from other files will be consolidated.

Consolidate File
This button is used to select other files containing comments to be incorporated to the Main Consolidate File.

Output Consolidate File
This button is used to select a location and file name for comment Consolidation.

Consolidate Comment
This button is used to execute the comment consolidation after Main Consolidate File, Consolidate File, Output Consolidate file are selected. It will bring comments from Consolidate File to the Main Consolidate File if the same record is identified.

Clean
This button is used to clear the previous setting/selection for Old Excel File
and New Excel File as well as Output Excel File.

Include/Exclude Comment
This button is used to define if comment in the old excel file would be transferred
to the comparison result for identical outputs after comparison.
This default as include comment

Auto Date Fill/No Date Fill
This button is used to define if date (yyyymmdd) will be automatically added to
comparison result file name.  i.e. Compared result_yyyymmdd.xls.
This default as auto date fill
This button is used to set maximum columns considered during comparison,
table split rule, comments pattern, and set default value for auto date fill
and include/exclude comment.

Help
This button is used to review the brief instruction, version and revision history.

Quit
This button is used to quit the program.
###############################################################################

Author: Yao Jian (yaojian2@merck.com)
Release: 3.0
GDO Comparator Tool
	Introduced new functionality to run Cumulative compare
	Introduced new functionality to consolidate comment
Excel Add-in
	Added short cut keys for quick comment
	Added new functionality to highlight brackets, certain phrase and compare content in selected area.
	Updated advance filter to allow additional filter after advance filter is executed.
Release: 2.0
    GDO Comparator Tool: 
        Updated Comments Pattern default to CDM Comment
    Excel Add-in
        Updated Naming convention rule to better Identify report ID
        Added DMC Reports Tools function to enable sorting tabs alphabetically, adding comments pattern column, filter blank in comment column, sort by color (green first), exclude countries from output, create summary tab
        Added Search Site/Subject function and enable filtering of searched value
        Added advanced search function
        Added function to search certain expression in workbook
        Added function to remove all filters
        Added function to clear tab color
        Added function to create a summary tab to list all tabs with short cut link for quick access to different tabs


Release: 1.20
   Updated DMCtools to enable consolidation when there is more than one tab/sheet in one excel
   Updated DMCtools to enable consolidation of InForm Cognos report and J Review reports
   Updated DMCtools to identify JReview report by looking for _JR in excel content, and the sheet name after consolidation will remain the same as in the original file. For reports without _JR in content, sheet name will be renamed using previous rule

Release: 1.15
   Update Table Split to new format;
   Update DMCTools support the Maximum Font Size Algorithm to choose SheetName.
   Update DMCTools support Long Name with special charactor to choose SheetName.
   Update the format of Cell to light_green as well as border
   Update to single output for new/old document only.

Release: 1.0
"""

import ConfigParser


def setConfig(MAX_COLUMNS=50, TABLE_SPLIT="^\s*[Tt]able\s*\d+\s*:?\s*$", COMMENT_PATTERN="CDM Comment",
              COMMENTS="Include Comments", AUTODATE="Auto Date Fill", filename="config.ini"):
    """Set Configures in config.ini"""
    with open(filename, "w") as INI:
        INI.write("""
[main]
MAX_COLUMNS = {}
TABLE_SPLIT = {}
COMMENT_PATTERN = {}
COMMENTS = {}
AUTODATE = {}
        """.format(MAX_COLUMNS, TABLE_SPLIT, COMMENT_PATTERN, COMMENTS, AUTODATE))


def getConfig(filename="config.ini"):
    """Get Configures
    """
    conf = ConfigParser.ConfigParser()
    conf.read(filename)
    MAX_COLUMNS = int(conf.get("main", "MAX_COLUMNS"))
    TABLE_SPLIT = conf.get("main", "TABLE_SPLIT")
    COMMENT_PATTERN = conf.get("main", "COMMENT_PATTERN")
    COMMENTS = conf.get("main", "COMMENTS")
    AUTODATE = conf.get("main", "AUTODATE")
    return MAX_COLUMNS, TABLE_SPLIT, COMMENT_PATTERN, COMMENTS, AUTODATE


class TkConfig(Tkinter.Frame):
    def __init__(self, root):
        """ init config frame
        """
        Tkinter.Frame.__init__(self, root)

        MAX_COLUMNS, TABLE_SPLIT, COMMENT_PATTERN, COMMENTS, AUTODATE = getConfig()

        frame1 = Tkinter.Frame(root)
        frame1.pack(side=Tkinter.TOP, fill="x")
        Tkinter.Label(frame1, text="Maximum Columns", width=20).pack(side=Tkinter.LEFT, padx=10, pady=10, fill="x")
        self.mVar = Tkinter.StringVar(self)
        self.mVar.set(MAX_COLUMNS)
        self.mE = Tkinter.Entry(frame1, width=20, textvariable=self.mVar)
        self.mE.pack(side=Tkinter.TOP, padx=10, pady=10, fill="x")

        frame2 = Tkinter.Frame(root)
        frame2.pack(side=Tkinter.TOP, fill="x")
        Tkinter.Label(frame2, text="Table Split", width=20).pack(side=Tkinter.LEFT, padx=10, pady=10)
        self.tVar = Tkinter.StringVar(self)
        self.tVar.set(TABLE_SPLIT)
        self.tE = Tkinter.Entry(frame2, width=20, textvariable=self.tVar)
        self.tE.pack(side=Tkinter.RIGHT, padx=10, pady=10)

        frame3 = Tkinter.Frame(root)
        frame3.pack(side=Tkinter.TOP, fill="x")
        Tkinter.Label(frame3, text="Comments Pattern", width=20).pack(side=Tkinter.LEFT, padx=10, pady=10)
        self.cVar = Tkinter.StringVar(self)
        self.cVar.set(COMMENT_PATTERN)
        self.cE = Tkinter.Entry(frame3, width=20, textvariable=self.cVar)
        self.cE.pack(side=Tkinter.RIGHT, padx=10, pady=10)

        frame4 = Tkinter.Frame(root)
        frame4.pack(side=Tkinter.TOP, fill="x")
        Tkinter.Label(frame4, text="Auto Date Fill", width=20).pack(side=Tkinter.LEFT, padx=10, pady=10)
        self.aVar = Tkinter.StringVar(self)
        self.aVar.set(AUTODATE)
        self.aE = Tkinter.OptionMenu(frame4, self.aVar, "Auto Date Fill", "No Date Fill")
        self.aE.pack(side=Tkinter.RIGHT, padx=10, pady=10)

        frame5 = Tkinter.Frame(root)
        frame5.pack(side=Tkinter.TOP, fill="x")
        Tkinter.Label(frame5, text="Comments Including", width=20).pack(side=Tkinter.LEFT, padx=10, pady=10)
        self.sVar = Tkinter.StringVar(self)
        self.sVar.set(COMMENTS)
        self.sE = Tkinter.OptionMenu(frame5, self.sVar, "Include Comments", "Exclude Comments")
        self.sE.pack(side=Tkinter.RIGHT, padx=10, pady=10)

        frame6 = Tkinter.Frame(root)
        frame6.pack(side=Tkinter.TOP, fill="x")
        Tkinter.Button(frame6, text="Cancel", command=self.cancel).pack(side=Tkinter.RIGHT, padx=10, pady=10)
        Tkinter.Button(frame6, text="Save", command=self.save).pack(side=Tkinter.RIGHT, padx=10, pady=10)

        Tkinter.Button(root, text="Restore Default Value", command=self.restore).pack(side=Tkinter.BOTTOM, padx=10,
                                                                                      pady=10)

    def save(self):
        """ save new configure
        """
        try:
            int(self.mVar.get())
        except:
            tkMessageBox.showerror("Error:", "Maximum Column should be a Integer!")
            return

        newconfig = (self.mVar.get(), self.tVar.get(), self.cVar.get(), self.sVar.get(), self.aVar.get())
        result = tkMessageBox.askquestion("Info", "Are you want to save new values and restart?", icon="info")
        if result == "yes":
            global CONFIG
            global MAIN
            setConfig(*newconfig)
            CONFIG.destroy()
            MAIN.destroy()
            main()

    def cancel(self):
        result = tkMessageBox.askquestion("Info", "Not Save modifications?", icon="info")
        if result == "yes":
            global CONFIG
            CONFIG.destroy()

    def restore(self):
        result = tkMessageBox.askquestion("Info", "Are you want to restore values of the system and restart?",
                                          icon="info")
        if result == "yes":
            global CONFIG
            global MAIN
            setConfig()
            CONFIG.destroy()
            MAIN.destroy()
            main()


class TkFile(Tkinter.Frame):
    """Main Frame of the tool"""

    def __init__(self, root):
        """ init the frame
        """
        Tkinter.Frame.__init__(self, root)
        # find EXCEL.EXE
        import os
        try:
            ins, outs = os.popen2("assoc .xlam")
            ftype = outs.readline().split("=")[1]
            ins, outs = os.popen2("ftype {}".format(ftype))
            self.excel = outs.readline().split('"')[1]
        except:
            self.errorMessage("You should install EXCEL 2007 or later in your system")
            return

        MAX_COLUMNS, TABLE_SPLIT, COMMENT_PATTERN, COMMENTS, AUTODATE = getConfig()
        autodate = AUTODATE  # get default auto date
        comments = COMMENTS  # get default comments

        button_opt = {'fill': Tkconstants.BOTH, 'padx': 35, 'pady': 5}

        self.o = Tkinter.Button(self, text='Old Excel File', command=self.old)
        self.o.pack(**button_opt)
        self.o.configure(bg="lightgrey")

        self.n = Tkinter.Button(self, text='New Excel File', command=self.new)
        self.n.pack(**button_opt)
        self.n.configure(bg="lightgrey")

        self.t = Tkinter.Button(self, text='Output Excel File', command=self.output)
        self.t.pack(**button_opt)
        self.t.configure(bg="lightgrey")

        self.r = Tkinter.Button(self, text='Run Compare', command=self.compare)
        self.r.pack(**button_opt)
        self.r.configure(bg="grey")

        self.p = Tkinter.Button(self, text='Cumulative Compare', command=self.cumulative_compare)
        self.p.pack(**button_opt)
        self.p.configure(bg="grey")

        self.m = Tkinter.Button(self, text='Main Consolidate File', command=self.consolidate_one)
        self.m.pack(**button_opt)
        self.m.configure(bg="lightgrey")

        self.k = Tkinter.Button(self, text='Consolidate File', command=self.consolidate_two)
        self.k.pack(**button_opt)
        self.k.configure(bg="lightgrey")

        self.g = Tkinter.Button(self, text='Output Consolidate File', command=self.output_con)
        self.g.pack(**button_opt)
        self.g.configure(bg="lightgrey")

        self.s = Tkinter.Button(self, text='Consolidate Comment', command=self.consolidate_comment)
        self.s.pack(**button_opt)
        self.s.configure(bg="grey")

        self.c = Tkinter.Button(self, text='Clean', command=self.clean)
        self.c.pack(**button_opt)
        self.c.configure(bg="lightgrey")

        self.commentsVar = Tkinter.StringVar(self)
        self.commentsVar.set(comments)
        self.c = Tkinter.OptionMenu(self, self.commentsVar, "Include Comments", "Exclude Comments")
        self.c.pack(**button_opt)
        self.c.configure(bg="lightgrey")

        self.autodateVar = Tkinter.StringVar(self)
        self.autodateVar.set(autodate)
        self.d = Tkinter.OptionMenu(self, self.autodateVar, "Auto Date Fill", "No Date Fill")
        self.d.pack(**button_opt)
        self.d.configure(bg="lightgrey")

        self.config = Tkinter.Button(self, text='Setting', command=self.setting)
        self.config.pack(**button_opt)
        self.config.configure(bg="lightgrey")

        self.h = Tkinter.Button(self, text='Help', command=self.help)
        self.h.pack(**button_opt)
        self.h.configure(bg="lightgrey")

        self.q = Tkinter.Button(self, text='Quit', command=MAIN.destroy)
        self.q.pack(**button_opt)
        self.q.configure(bg='lightgrey')

        self.oldD = self.newD = self.outputD = ""  # set NULL value to the directory
        self.consolidate_newD = self.con_oldD = self.outputCon = ""

        self.dir_opt = options = {}
        options['initialdir'] = 'C:\\'
        options['filetype'] = (("Excel 97-2003 Format", "*.xls"),)
        options['parent'] = root

        # install DMCTools to Excel
        try:
            USERNAME = os.getenv("USERNAME")
        except:
            self.errorMessage("Missing User Name")
            return

        self.XLStart = USR_XLSTART.format(USERNAME=USERNAME)
        if not os.path.exists(self.XLStart):
            try:
                os.mkdir(self.XLStart)
            except:
                self.errorMessage("Cannot make directory: {}".format(self.XLStart))
                return

        if not os.path.exists(os.path.join(self.XLStart, DMCTools)):
            self.Consolidation()

    def Consolidation(self):
        """ install Consolidation """
        import os
        import glob
        import sys
        # tkMessageBox.showinfo("Info", "Install DMCTOOL in to your EXCEL,\n Close Excel First!")

        for filename in glob.glob(self.XLStart + "\\DMCTools*.xlam"):
            stdin, stdout, stderr = os.popen3('DEL /F "{}" '.format(filename))
            error = stderr.read()
            if len(error) > 0:
                # tkMessageBox.showerror("Error", "Close Excel First! Run again!")
                sys.exit()
                return
        ins, outs = os.popen2('copy {} "{}" '.format(DMCTools, self.XLStart))

    def errorMessage(self, txt):
        """ show errorMessage
        """
        tkMessageBox.showerror("Error:", txt)

    def clean(self):
        """ Clean all the directory values to NULL
        """
        if self.oldD != "" or self.newD != "":
            result = tkMessageBox.askquestion("Run Compare Info:",
                                              "Cleaned all Directory?\n\nOld :\n{}\n\nNew :\n{}\n\nOutput :\n{}\n\n".format(
                                                  self.oldD, self.newD, self.outputD), icon='info')
            if result == 'yes':
                self.oldD = self.newD = self.outputD = ""
                self.o.configure(bg="lightgrey")
                self.n.configure(bg="lightgrey")
                self.t.configure(bg="lightgrey")

        if self.consolidate_newD != "" or len(self.con_oldD) != 0:
            self.consolidate_oldD = str("\n".join(self.con_oldD))
            result = tkMessageBox.askquestion("Cumulative Compare Info :",
                                              "Cleaned all Directory?\n\nMain Consolidate File :\n{}\n\nConsolidate File :\n{}\n\nConsolidate Output File :\n{}".format(
                                                  self.consolidate_newD, self.consolidate_oldD, self.outputCon),
                                              icon='info')
            if result == 'yes':
                self.consolidate_newD = ""
                self.con_oldD = ""
                self.outputCon = ""
                self.m.configure(bg="lightgrey")
                self.k.configure(bg="lightgrey")
                self.g.configure(bg="lightgrey")

    def help(self):
        """ Show Help Information
        """
        from Tkinter import Text, INSERT
        newroot = Tkinter.Tk()
        newroot.title("Help")
        text = Text(newroot)
        text.insert(INSERT, Help)
        text.pack()

    def setting(self):
        """ Setting Configurations
        """
        global CONFIG
        CONFIG = Tkinter.Tk()
        CONFIG.title("Default Value Configuration")
        TkConfig(CONFIG).pack()
        CONFIG.mainloop()

    def compare(self, cumulative_compare=False):
        """ Main comparing method
        """
        txt = []
        if self.oldD == "":
            txt.append("Old Excel File is Missing!")
        if self.newD == "":
            txt.append("New Excel File is Missing!")
        if self.outputD == "":
            txt.append("Output Excel File is Missing!")

        txt = "\n".join(txt)
        if txt != "":
            self.errorMessage(txt)
            return
        else:
            if self.outputD.lower().endswith(".xls"):
                this = self.outputD[:-4]
            else:
                this = self.outputD

            if self.autodateVar.get() == "Auto Date Fill":
                import time
                tag = "_" + time.strftime("%Y%m%d")
                if not this.endswith(tag):
                    this += tag

            this += ".xls"
            self.outputD = this

            result = tkMessageBox.askquestion("Info:",
                                              "Old :\n{}\n\nNew :\n{}\n\nOutput :\n{}\n\n{}\n\nDo you want to run?".format(
                                                  self.oldD, self.newD, self.outputD, self.commentsVar.get()),
                                              icon='info')
            if result != 'yes':
                return

        newroot = None
        try:  # run compare method
            from comparing import compare
            from Tkinter import Text, INSERT
            newroot = Tkinter.Tk()
            newroot.lift()
            newroot.title("Processing")
            text = Text(newroot)  # Output the processing information
            text.pack()

            MAX_COLUMNS, TABLE_SPLIT, COMMENT_PATTERN, COMMENTS, AUTODATE = getConfig()

            if self.commentsVar.get() == "Include Comments":
                comments = True
            else:
                comments = False

            for txt in compare(self.oldD, self.newD, self.outputD, cumulative_compare, comments, COMMENT_PATTERN,
                               TABLE_SPLIT, MAX_COLUMNS):
                text.insert(INSERT, txt + "\n")
                if txt.startswith("Warning"):
                    tkMessageBox.showwarning("Warning", txt[8:])
        except Exception as e:
            self.errorMessage(e.message)
            newroot.destroy()
            return

        result = tkMessageBox.askquestion("Info", "Comparing finished! \n Open the OUTPUT file?", icon="info")
        #  newroot.destroy()
        if result == "yes":
            import subprocess
            subprocess.Popen([self.excel, self.outputD])

    def consolidate_comment(self):
        txt = []
        if self.consolidate_newD == "":
            txt.append("Main Consolidate File is Missing!")
        if len(self.con_oldD) == 0:
            txt.append("Consolidate File is Missing!")
        if self.outputCon == "":
            txt.append("Consolidate Output File is Missing!")
        txt = "\n".join(txt)
        if txt != "":
            self.errorMessage(txt)
            return
        else:
            if self.outputCon.lower().endswith(".xls"):
                this = self.outputCon[:-4]
            else:
                this = self.outputCon
            if self.autodateVar.get() == "Auto Date Fill":
                import time
                tag = "_" + time.strftime("%Y%m%d")
                if not this.endswith(tag):
                    this += tag

            this += ".xls"
            self.outputCon = this
            self.consolidate_oldD = str("\n".join(self.con_oldD))
            result = tkMessageBox.askquestion("Info:",
                                              "Main Consolidate File:\n{}\n\nConsolidate File:\n{}\n\nConsolidate Output:\n{}\n\n{}\n\nDo you want to run?".format(
                                                  self.consolidate_newD, self.consolidate_oldD,
                                                  self.outputCon,
                                                  self.commentsVar.get()),
                                              icon='info')
            if result != 'yes':
                return

        newroot = None
        try:
            from comparing import consolidate_compare
            from Tkinter import Text, INSERT
            newroot = Tkinter.Tk()
            newroot.lift()
            newroot.title("Processing")
            text = Text(newroot)  # Output the processing information
            text.pack()
            MAX_COLUMNS, TABLE_SPLIT, COMMENT_PATTERN, COMMENTS, AUTODATE = getConfig()
            for txt in consolidate_compare(self.con_oldD, self.consolidate_newD, self.outputCon, COMMENT_PATTERN,
                                           TABLE_SPLIT):
                text.insert(INSERT, txt + "\n")
                if txt.startswith("Warning"):
                    tkMessageBox.showwarning("Warning", txt[8:])
        except Exception as e:
            self.errorMessage(e.message)
            newroot.destroy()
            return

        result = tkMessageBox.askquestion("Info", "Consolidate finished! \n Open the consolidate output file ?",
                                          icon="info")
        if result == "yes":
            import subprocess
            subprocess.Popen([self.excel, self.outputCon])

    # 进行累计比较
    def cumulative_compare(self):
        cumulative_compare = True
        self.compare(cumulative_compare)

    def old(self):
        """ Add Old Directory
        """
        self.dir_opt['initialdir'] = 'C:\\TEMP\\EXAMPLE\\OUTPUT'
        self.dir_opt['title'] = 'Open Old File'
        self.oldD = tkFileDialog.askopenfilename(**self.dir_opt)

        try:
            self.oldD.encode("ascii")
        except:
            self.errorMessage("Filename/Directory include Non-English character is not support!")
            return

        if self.oldD != "":
            if not self.oldD.lower().endswith(".xls"):
                self.oldD = ""
                self.errorMessage("Only xls format are acceptable.")
            else:
                with open(self.oldD) as XLS:
                    txt = XLS.readline()
                    if 'html' in txt.lower():
                        self.errorMessage(
                            "This xls format is not true xls. \n Please use consolidation to process it first!")
                        self.oldD = ""
                        return
                    else:
                        self.o.configure(bg="blue")

    def new(self):
        """ Add New Directory
        """
        self.dir_opt['initialdir'] = 'C:\\TEMP\\EXAMPLE\\OUTPUT'
        self.dir_opt['title'] = 'Open New File'
        self.newD = tkFileDialog.askopenfilename(**self.dir_opt)

        try:
            self.newD.encode("ascii")
        except:
            self.errorMessage("Filename/Directory include Non-English character is not support!")
            return

        if self.newD != "":
            if not self.newD.lower().endswith(".xls"):
                self.newD = ""
                self.errorMessage("Only xls format are acceptable.")
            else:
                with open(self.newD) as XLS:
                    txt = XLS.readline()
                    if 'html' in txt.lower():
                        self.errorMessage(
                            "This xls format is not true xls. \n Please use consolidation to process it first!")
                        self.newD = ""
                        return
                    else:
                        result = tkMessageBox.askquestion("Warning",
                                                          "Please ensure the New .xls should have no GREEN color annotated, \nOR the results will be not correct.\n\n Are you sure?",
                                                          icon="info")
                        if result == "yes":
                            self.n.configure(bg="blue")
                        else:
                            import subprocess
                            subprocess.Popen([self.excel, self.newD])
                            self.newD = ""

    # new add
    def consolidate_one(self):
        """main consolidate file
        """
        self.dir_opt['initialdir'] = 'C:\\TEMP\\EXAMPLE\\OUTPUT'
        self.dir_opt['title'] = 'Main Consolidate File'
        self.consolidate_newD = tkFileDialog.askopenfilename(**self.dir_opt)

        try:
            self.consolidate_newD.encode("ascii")
        except:
            self.errorMessage("Filename/Directory include Non-English character is not support!")
            return

        if self.consolidate_newD != "":
            if not self.consolidate_newD.lower().endswith(".xls"):
                self.consolidate_newD = ""
                self.errorMessage("Only xls format are acceptable.")
            else:
                with open(self.consolidate_newD) as XLS:
                    txt = XLS.readline()
                    if 'html' in txt.lower():
                        self.errorMessage(
                            "This xls format is not true xls. \n Please use consolidation to process it first!")
                        self.consolidate_newD = ""
                        return
                    else:
                        result = tkMessageBox.askquestion("Warning",
                                                          "Make sure Main Consolidate File and attach Consolidate File are the same except Comment\nOtherwise the results will be not correct.\n\n Are you sure?",
                                                          icon="info")
                        if result == "yes":
                            self.m.configure(bg="blue")
                        else:
                            import subprocess
                            subprocess.Popen([self.excel, self.consolidate_newD])
                            self.consolidate_newD = ""

    def consolidate_two(self):
        """consolidate file in folder
        """
        self.con_oldD = []
        self.dir_opt['initialdir'] = 'C:\\TEMP\\EXAMPLE\\OUTPUT'
        self.dir_opt['title'] = 'Consolidate file in folder'
        self.dir_opt['filetypes'] = [("Excel 97-2003 Format", "*.xls")]
        master = Tkinter.Tk()
        master.withdraw()  # 不显示界面主窗口
        self.fnstr = tkFileDialog.askopenfilenames(**self.dir_opt)
        self.fns = master.tk.splitlist(self.fnstr)  # 把多个文件名字符串分割成元组
        for i in range(len(self.fns)):
            try:
                self.fns[i].encode("ascii")
            except:
                self.errorMessage("Filename/Directory include Non-English character is not support!")
                return

            if self.fns[i] != "":
                with open(self.fns[i]) as XLS:
                    txt = XLS.readline()
                    if 'html' in txt.lower():
                        self.errorMessage("This xls format is not true xls.")
                        self.fns[i] = ""
                        return
                    else:
                        self.k.configure(bg="blue")
                        self.con_oldD.append(self.fns[i])

    def output(self):
        """ Add Output Directory
        """
        self.dir_opt['initialdir'] = 'C:\\TEMP\\EXAMPLE\\OUTPUT'
        self.dir_opt['title'] = 'Open Output File'
        self.outputD = tkFileDialog.asksaveasfilename(**self.dir_opt)

        try:
            self.outputD.encode("ascii")
        except:
            self.errorMessage("Filename/Directory include Non-English character is not support!")
            return

        if self.outputD != "":
            self.t.configure(bg="blue")

    def output_con(self):
        """ Add Consolidate Output Directory
        """
        self.dir_opt['initialdir'] = 'C:\\TEMP\\EXAMPLE\\OUTPUT'
        self.dir_opt['title'] = 'Open Consolidate Output File'
        self.outputCon = tkFileDialog.asksaveasfilename(**self.dir_opt)

        try:
            self.outputCon.encode("ascii")
        except:
            self.errorMessage("Filename/Directory include Non-English character is not support!")
            return

        if self.outputCon != "":
            self.g.configure(bg="blue")


def main():
    'main function'
    global MAIN
    MAIN = Tkinter.Tk()
    MAIN.title("GDOcom")
    TkFile(MAIN).pack()
    MAIN.mainloop()


MAIN = None
CONFIG = None

if __name__ == "__main__":
    main()
