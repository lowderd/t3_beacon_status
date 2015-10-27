__author__ = 'Duncan Lowder'
__version__ = '0.1'

# -----------------------------------------------------------------------------
#
# get_beacon_status.py
#
# Description:
#   This script obtains all of the information related to a beacon from the
#   manufacturing database and displays it.
#
# Building the Application
#   To build an executable and installer for this application, the following
#   steps must be completed:
#       1. Verify that the script runs correctly
#           Increment the __version__ # and run the script and any available
#           tests.
#           TO DO: Implement a test library to verify the functionality
#                  of the script
#       2. Build an executable by running py2exe
#           Open a command prompt in the directory of the script and check
#           that the setup.py script for get_beacon_status.py is in the same
#           directory. From the command prompt, execute the command:
#           "python setup.py py2exe"
#       3. Build an installer
#           Open the "installer_script.iss" file, which is located within the
#           "dist" folder, with InnoSetup. Compile the script and run the
#           installer to verify that the setup.exe package works correctly.
#           The setup.exe package is located within the Output folder in dist.
#
# Packages and Resources:
#   wxPython  - http://www.wxpython.org/
#   pyodbc    - https://github.com/mkleehammer/pyodbc
#   py2exe    - http://www.py2exe.org/
#   InnoSetup - http://www.jrsoftware.org/isinfo.php
#
# -----------------------------------------------------------------------------

# -----------------------------------------------------------------------------
# IMPORTS
# -----------------------------------------------------------------------------

import wx
import wx.lib
import wx.lib.flatnotebook as fnb
import wx.grid

import pyodbc
import logging
import os

import json

from operator import itemgetter

# -----------------------------------------------------------------------------
# WORKING DIRECTORY
# -----------------------------------------------------------------------------

# Set working directory
try:
    abspath = os.path.abspath(__file__)
    dname = os.path.dirname(abspath)
    os.chdir(dname)
except NameError:
    import sys
    os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))


# -----------------------------------------------------------------------------
# LOGGING SETUP
# -----------------------------------------------------------------------------

# Create logger
logger = logging.getLogger("beacon_status")
logger.setLevel(logging.DEBUG)

# Create log file handler
# log_fh = logging.FileHandler("beacon_status.log")
# log_fh.setLevel(logging.ERROR)

# Create console handler with higher level
log_ch = logging.StreamHandler()
log_ch.setLevel(logging.ERROR)

# Create formatter and add it to the handlers
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
formatter2 = logging.Formatter("%(name)-14s: %(levelname)-8s %(message)s")
# log_fh.setFormatter(formatter)
log_ch.setFormatter(formatter2)

# Add the handlers to the logger
# logger.addHandler(log_fh)
logger.addHandler(log_ch)

# -----------------------------------------------------------------------------
# GLOBAL VARIABLES
# -----------------------------------------------------------------------------

db_table_list = ["assemblyKittingTable", "DFTestingTable", "calibrationTable", "finalTestTable", "closeCaseTable",
                 "finalInspectionTable", "packagingTable", "falloutTable", "engineeringReworkTable"]

APP_EXIT = 1
NEW_QUERY = 2
SAVE_RESULTS = 3
ABOUT_BOX = 4
OPEN_FILE = 5
QUICK_HELP = 6

# Directory to the applications icon
# app_icon = "icons\\BCALogoMedium.png"
# app_icon = "icons\\app_icon_radar.png"
app_icon = "icons\\app_icon_radar.ico"

# SQL database connection string. The BCAUser account allows read-only access to the T3Production database
sql_cnxn_str = "DRIVER={SQL Server};SERVER=172.18.149.5,2222;DATABASE=T3Production;UID=BCAUser;PWD=*trekkie#123;" \
               "Trusted_Connection=no"


# -----------------------------------------------------------------------------
# CLASSES
# -----------------------------------------------------------------------------

class CustomTaskBarIcon(wx.TaskBarIcon):

    def __init__(self, parent):
        super(CustomTaskBarIcon, self).__init__()

        self.parent = parent

        # Setup
        ico = wx.Icon(app_icon, wx.BITMAP_TYPE_PNG)
        self.SetIcon(ico)
        self.parent.SetIcon(ico)

        # Event Handlers
        self.Bind(wx.EVT_MENU, self.on_menu)

    def create_popup_menu(self):
        """
            Base class virtual method for creating the popup menu for the icon
        """
        menu = wx.Menu()
        menu.Append(NEW_QUERY, "New Query")
        menu.Append(SAVE_RESULTS, "Save Results")
        menu.AppendSeparator()
        menu.Append(APP_EXIT, "Exit")

        return menu

    def on_menu(self, e):
        evt_id = e.GetId()

        if evt_id == NEW_QUERY:
            self.parent.new_query()
        elif evt_id == SAVE_RESULTS:
            self.parent.save_results()
        elif evt_id == APP_EXIT:
            self.parent.on_quit()
        else:
            e.Skip()


class FileMenu(wx.Menu):

    def __init__(self, parent):
        super(FileMenu, self).__init__()

        new_query_item = wx.MenuItem(self, NEW_QUERY, "&New Query\tCtrl+N")
        new_query_item.SetBitmap(wx.Bitmap("icons\search25.png"))

        open_file_item = wx.MenuItem(self, OPEN_FILE, "&Open Report File\tCtrl+O")
        open_file_item.SetBitmap(wx.Bitmap("icons\\add25.png"))

        save_results_item = wx.MenuItem(self, SAVE_RESULTS, "&Save Results\tCtrl+S")
        save_results_item.SetBitmap(wx.Bitmap("icons\down25.png"))

        quit_item = wx.MenuItem(self, APP_EXIT, "&Quit\tCtrl+Q")
        quit_item.SetBitmap(wx.Bitmap("icons\close25.png"))

        # Append Menu Items
        self.AppendItem(new_query_item)
        self.AppendItem(open_file_item)
        self.AppendItem(save_results_item)
        self.AppendSeparator()
        self.AppendItem(quit_item)

        # Bind Menu Items
        self.Bind(wx.EVT_MENU, parent.new_query, new_query_item)
        self.Bind(wx.EVT_MENU, parent.save_results, save_results_item)
        self.Bind(wx.EVT_MENU, parent.on_quit, quit_item)


class ViewMenu(wx.Menu):

    def __init__(self, parent):
        super(ViewMenu, self).__init__()

        self.shst = self.Append(wx.ID_ANY, "Show Status", "Show Status", kind=wx.ITEM_CHECK)
        self.shtl = self.Append(wx.ID_ANY, "Show Toolbar", "Show Toolbar", kind=wx.ITEM_CHECK)

        self.Check(self.shst.GetId(), True)
        self.Check(self.shtl.GetId(), True)

        self.Bind(wx.EVT_MENU, parent.toggle_status_bar, self.shst)
        self.Bind(wx.EVT_MENU, parent.toggle_tool_bar, self.shtl)


class ResultsNotebook(fnb.FlatNotebook):

    def __init__(self, parent):
        super(ResultsNotebook, self).__init__(parent)

        # Notebook attributes
        self.empty_page = wx.Panel(self)
        self.empty_page.SetBackgroundColour("#6C727F")

        # Setup
        self.AddPage(self.empty_page, "< ... >")


class SerialNumberDialog(wx.Dialog):

    def __init__(self, parent):
        super(SerialNumberDialog, self).__init__(parent)

        self.init_ui()
        self.SetSize((300, 160))
        self.SetTitle("Enter Serial Number")

    def init_ui(self):

        pnl = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # Main panel for serial number input
        sb = wx.StaticBox(pnl, label="Beacon Serial Number")
        sbs = wx.StaticBoxSizer(sb, orient=wx.VERTICAL)
        sbs.Add(wx.StaticText(pnl, label="Enter or scan the top-level serial number of the unit."))
        sbs.AddSpacer(10)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox1.Add(wx.StaticText(pnl, label="Serial Number:"), flag=wx.ALIGN_CENTER)

        self.sn_text = wx.TextCtrl(pnl)
        hbox1.Add(self.sn_text, flag=wx.LEFT, border=10)

        sbs.Add(hbox1, flag=wx.ALIGN_CENTER)

        pnl.SetSizer(sbs)

        # Ok and cancel button panel/sizer
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        ok_button = wx.Button(self, label="Ok")
        cancel_button = wx.Button(self, label="Cancel")
        hbox2.Add(ok_button)
        hbox2.Add(cancel_button, flag=wx.LEFT, border=5)

        # Arrange serial number input and button panels
        vbox.Add(pnl, proportion=1, flag=wx.ALL | wx.EXPAND, border=5)
        vbox.Add(hbox2, flag=wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, border=10)

        self.SetSizer(vbox)

        ok_button.Bind(wx.EVT_BUTTON, self.format_sn)
        cancel_button.Bind(wx.EVT_BUTTON, self.on_close)

    def format_sn(self, e):
        logger.debug("SerialNumberDialog:format_sn")
        self.serial_number = self.sn_text.GetValue().upper()
        logger.debug("SerialNumberDialog:format_sn -> serial number: {0}".format(self.serial_number))

        self.Destroy()

    def on_close(self, e):
        logger.debug("SerialNumberDialog:on_close")
        self.Destroy()


class DfTable(wx.grid.Grid):

    def __init__(self, parent, beacon_data):
        super(DfTable, self).__init__(parent)

        self.CreateGrid(1, 7)

        # Set column values
        col_vals = ["VL", "AL", "VX", "AX", "VY", "AY", "VN"]
        for i in range(len(col_vals)):
            self.SetColLabelValue(i, col_vals[i])
            self.SetRowLabelValue(0, "DF Values:")

            try:
                self.SetCellValue(0, i, str(beacon_data[col_vals[i]]).upper())
            except KeyError:
                self.SetCellValue(0, i, "N/A")

        self.AutoSize()


class ResultsPage(wx.Panel):

    def __init__(self, parent, beacon_data, serial_number):
        super(ResultsPage, self).__init__(parent)

        pnl1 = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        sb1 = wx.StaticBox(self, label="Beacon Information")
        sb1s = wx.StaticBoxSizer(sb1, orient=wx.VERTICAL)

        # Set Beacon Information string, display N/A for transaction time if no
        # information was found
        if len(beacon_data) is not 0:
            header_str = "Serial Number: {0}, First Scanned: {1}".format(serial_number,
                                                                         beacon_data[0]["transactionTime"])
        else:
            header_str = "Serial Number: {0}, First Scanned: N/A".format(serial_number)

        sb1s.Add(wx.StaticText(self, label=header_str), flag=wx.LEFT)

        # Generate DF data table
        for table_entry in beacon_data:
            if table_entry["db_table"] == "DFTestingTable":
                df_table = DfTable(self, table_entry)
                sb1s.AddSpacer(5)
                sb1s.Add(df_table)

        pnl1.SetSizer(sb1s)

        sb2 = wx.StaticBox(self, label="Manufacturing Information")
        sb2s = wx.StaticBoxSizer(sb2, orient=wx.VERTICAL)

        # Check if no information was found
        if len(beacon_data) is 0:
            sb2s.Add(wx.StaticText(self, label="No data found"), flag=wx.LEFT, border=10)
        else:
            # Populate the manufacturing data entries
            results_tree = wx.TreeCtrl(self, style=wx.TR_DEFAULT_STYLE | wx.TR_HIDE_ROOT | wx.TR_TWIST_BUTTONS)
            root = results_tree.AddRoot("Mfg Data")

            for entry in beacon_data:
                logger.debug("ResultsPage: add entry {0}".format(entry))
                print entry["db_table"]
                db_table_str = format_db_table_str(entry["db_table"])

                # Add DB tables
                table_entry = results_tree.AppendItem(root, "{0}: {1}".format(entry["transactionTime"], db_table_str))

                # Add all of the keys for the selected DB table
                for key in sorted(entry):
                    if key in ["db_table", "transactionID", "transactionTime", "serialNumber"]:
                        pass
                    elif key == "employeeID":
                        # Retrieve Employee Name to display
                        logger.debug("ResultsPage: key={0}, data={1}".format(key, entry[key]))
                        temp_key = results_tree.AppendItem(table_entry, format_column_str(key))
                        results_tree.AppendItem(temp_key, get_employee_name(entry[key]))
                    elif key == "failureCode":
                        if entry["failureCode"] == 0 or entry["failureCode"] is None:
                            pass
                        else:
                            logger.debug("ResultsPage: key={0}, data={1}".format(key, entry[key]))
                            temp_key = results_tree.AppendItem(table_entry, format_column_str(key))
                            results_tree.AppendItem(temp_key, "{0}: {1}".format(str(entry[key]),
                                                                                get_failure_description(entry[key])))
                            results_tree.ExpandAllChildren(temp_key)
                            results_tree.Expand(table_entry)
                    elif key == "failureDescription":
                        if entry["failureDescription"] == "Pass" or entry["failureDescription"] is None:
                            pass
                        else:
                            logger.debug("ResultsPage: key={0}, data={1}".format(key, entry[key]))

                            # Format failureDescription String
                            # Multiple strings can be entered, so these are split and then multiple entries
                            # are made within the Failure Description Page.
                            failure_str = str(entry[key]).split("\r\n")

                            temp_key = results_tree.AppendItem(table_entry, format_column_str(key))

                            for failure in failure_str:
                                results_tree.AppendItem(temp_key, failure)
                            results_tree.ExpandAllChildren(temp_key)
                            results_tree.Expand(table_entry)
                    else:
                        logger.debug("ResultsPage: key={0}, data={1}".format(key, entry[key]))
                        temp_key = results_tree.AppendItem(table_entry, format_column_str(key))
                        results_tree.AppendItem(temp_key, str(entry[key]))

            logger.debug("ResultsPage: quick best size = {0}".format(results_tree.GetQuickBestSize()))
            results_tree.SetQuickBestSize(results_tree.GetQuickBestSize())
            sb2s.Add(results_tree, 1, wx.EXPAND)
        sb2s.RecalcSizes()

        vbox.Add(pnl1, flag=wx.ALL | wx.EXPAND)
        vbox.AddSpacer(10)
        vbox.Add(sb2s, proportion=1, flag=wx.ALIGN_CENTER | wx.EXPAND)

        self.SetSizer(vbox)


class HelpDialog(wx.Dialog):

    def __init__(self, parent):
        super(HelpDialog, self).__init__(parent)

        h1_font = wx.Font(18, wx.SWISS, wx.SLANT, wx.NORMAL)
        h2_font = wx.Font(13, wx.SWISS, wx.NORMAL, wx.NORMAL)
        h3_font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL)

        self.SetTitle("BCA Tracker 3 Beacon Tracker Quick Start")
        self.SetIcon(wx.Icon(app_icon))

        pnl = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(pnl, wx.ID_ANY, "BCA Tracker 3 Beacon Tracker")
        title.SetFont(h1_font)

        info_str = "The BCA Tracker 3 Beacon Tracker application displays the recorded manufacturing information for a " \
                   "Tracker 3 Beacon."
        info = wx.StaticText(pnl, wx.ID_ANY, info_str)
        info.Wrap(400)

        instr = wx.StaticText(pnl, wx.ID_ANY, "Application Controls")

        instr.SetFont(h2_font)

        hb1 = wx.BoxSizer(wx.HORIZONTAL)
        hb1_ico = wx.StaticBitmap(pnl, bitmap=wx.BitmapFromImage(wx.Image("icons\search35.png", wx.BITMAP_TYPE_PNG)))
        hb1_txt1 = wx.StaticText(pnl, wx.ID_ANY, "New Query")
        hb1_txt1.SetFont(h3_font)
        hb1_txt2 = wx.StaticText(pnl, wx.ID_ANY, "Opens a dialog box which allows a Beacons serial number to be input.")
        hb1_txt2.Wrap(220)
        hb1.Add(hb1_ico, proportion=0.2, flag=wx.LEFT | wx.ALIGN_CENTER | wx.EXPAND)
        hb1.Add(hb1_txt1, flag=wx.ALIGN_CENTER | wx.ALL | wx.EXPAND, border=10)
        hb1.Add(hb1_txt2, flag=wx.LEFT, border=34)

        hb2 = wx.BoxSizer(wx.HORIZONTAL)
        hb2_ico = wx.StaticBitmap(pnl, bitmap=wx.BitmapFromImage(wx.Image("icons\\add35.png", wx.BITMAP_TYPE_PNG)))
        hb2_txt1 = wx.StaticText(pnl, wx.ID_ANY, "Open Report File")
        hb2_txt1.SetFont(h3_font)
        hb2_txt2 = wx.StaticText(pnl, wx.ID_ANY, "Opens a file dialog which allows a JSON Report file to be opened. The"
                                                 "opened file will be displayed in a new tab.")
        hb2_txt2.Wrap(220)
        hb2.Add(hb2_ico, proportion=0.2, flag=wx.LEFT | wx.ALIGN_CENTER | wx.EXPAND)
        hb2.Add(hb2_txt1, flag=wx.ALIGN_CENTER | wx.ALL | wx.EXPAND, border=10)
        hb2.Add(hb2_txt2, border=50)

        hb3 = wx.BoxSizer(wx.HORIZONTAL)
        hb3_ico = wx.StaticBitmap(pnl, bitmap=wx.BitmapFromImage(wx.Image("icons\down35.png", wx.BITMAP_TYPE_PNG)))
        hb3_txt1 = wx.StaticText(pnl, wx.ID_ANY, "Save Report File")
        hb3_txt1.SetFont(h3_font)
        hb3_txt2 = wx.StaticText(pnl, wx.ID_ANY, "Saves currently selected tab as a JSON Report file.")
        hb3_txt2.Wrap(220)
        hb3.Add(hb3_ico, proportion=0.2, flag=wx.LEFT | wx.ALIGN_CENTER | wx.EXPAND)
        hb3.Add(hb3_txt1, flag=wx.ALIGN_CENTER | wx.ALL | wx.EXPAND, border=10)
        hb3.Add(hb3_txt2, flag=wx.LEFT, border=3)

        hb4 = wx.BoxSizer(wx.HORIZONTAL)
        hb4_ico = wx.StaticBitmap(pnl, bitmap=wx.BitmapFromImage(wx.Image("icons\close35.png", wx.BITMAP_TYPE_PNG)))
        hb4_txt1 = wx.StaticText(pnl, wx.ID_ANY, "Exit Application")
        hb4_txt1.SetFont(h3_font)
        hb4_txt2 = wx.StaticText(pnl, wx.ID_ANY, "Closes all open tabs and exits the application.")
        hb4_txt2.Wrap(220)
        hb4.Add(hb4_ico, proportion=0.2, flag=wx.LEFT | wx.ALIGN_CENTER | wx.EXPAND)
        hb4.Add(hb4_txt1, flag=wx.ALIGN_CENTER | wx.ALL | wx.EXPAND, border=10)
        hb4.Add(hb4_txt2, flag=wx.LEFT, border=6)

        vbox.Add(title, flag=wx.ALL | wx.EXPAND | wx.ALIGN_CENTER, border=10)
        vbox.Add(info, flag=wx.LEFT | wx.EXPAND, border=10)
        vbox.Add(instr, flag=wx.ALL | wx.EXPAND, border=10)
        vbox.Add(hb1, flag=wx.LEFT | wx.BOTTOM | wx.EXPAND, border=10)
        vbox.Add(hb2, flag=wx.LEFT | wx.BOTTOM | wx.EXPAND, border=10)
        vbox.Add(hb3, flag=wx.LEFT | wx.BOTTOM | wx.EXPAND, border=10)
        vbox.Add(hb4, flag=wx.LEFT | wx.EXPAND, border=10)

        pnl.SetSizer(vbox)

        self.SetSize((400, 350))


class MainWindow(wx.Frame):
    """
        This class displays the main GUI for the BCA Beacon Tracker application
    """

    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)

        # Set Icon
        ico = wx.Icon(app_icon)
        self.SetIcon(ico)

        # Setup Menubar
        menu_bar = wx.MenuBar()
        menu_bar.SetBackgroundColour("#484537")
        self.file_menu = FileMenu(self)
        self.view_menu = ViewMenu(self)
        self.help_menu = wx.Menu()

        self.help_menu.Append(QUICK_HELP, "&Quick Start Guide")
        self.help_menu.AppendSeparator()
        self.help_menu.Append(ABOUT_BOX, "&About")
        self.Bind(wx.EVT_MENU, self.on_about_box, id=ABOUT_BOX)
        self.Bind(wx.EVT_MENU, self.on_help_box, id=QUICK_HELP)

        menu_bar.Append(self.file_menu, "&File")
        menu_bar.Append(self.view_menu, "&View")
        menu_bar.Append(self.help_menu, "&Help")

        self.SetMenuBar(menu_bar)

        self.toolbar = self.CreateToolBar()
        self.new_query_tool = self.toolbar.AddLabelTool(NEW_QUERY, "New Query", wx.Bitmap("icons\search35.png"),
                                                        shortHelp="New Query")
        self.open_file_tool = self.toolbar.AddLabelTool(OPEN_FILE, "Open Report File", wx.Bitmap("icons\\add35.png"),
                                                        shortHelp="Open Report File")
        self.save_results_tool = self.toolbar.AddLabelTool(SAVE_RESULTS, "Save Results", wx.Bitmap("icons\down35.png"),
                                                           shortHelp="Save Results")
        self.toolbar.AddSeparator()
        self.quit_tool = self.toolbar.AddLabelTool(APP_EXIT, "Exit", wx.Bitmap("icons\close35.png"), shortHelp="Exit")
        self.toolbar.SetBackgroundColour("#558AFC")
        self.toolbar.Realize()

        self.Bind(wx.EVT_MENU, self.new_query, self.new_query_tool)
        self.Bind(wx.EVT_MENU, self.save_results, self.save_results_tool)
        self.Bind(wx.EVT_MENU, self.open_file, self.open_file_tool)
        self.Bind(wx.EVT_MENU, self.on_quit, self.quit_tool)

        self.results_notebook = ResultsNotebook(self)
        self.page_counter = 0

        self.statusbar = self.CreateStatusBar()
        self.statusbar.SetStatusText("Ready...")

        self.SetSize((600, 500))
        self.SetTitle("BCA Tracker 3 Beacon Tracker")

        self.Center()
        self.Show()

    def toggle_status_bar(self, e):
        if self.view_menu.shst.IsChecked():
            self.statusbar.Show()
            log_str = "Show"
        else:
            self.statusbar.Hide()
            log_str = "Hide"
        logger.info("MainWindow:toggle_status_bar: {0}".format(log_str))

    def toggle_tool_bar(self, e):
        if self.view_menu.shtl.IsChecked():
            self.toolbar.Show()
            log_str = "Show"
        else:
            self.toolbar.Hide()
            log_str = "Hide"
        logger.info("MainWindow:toggle_tool_bar: {0}".format(log_str))

    def new_query(self, e):
        """
            This function displays a SerialNumberDialog window. After a serial
            number has been input, a query is then run on the T3Production
            database and the relevant information for the specified beacon is
            returned.
        :param e: Event ID
        :return: List of dictionaries containing manufacturing information for the specified beacon.
        """
        logger.debug("MainWindow:new_query")
        self.statusbar.SetStatusText('Waiting for Serial Number input...')

        ser_num_diag = SerialNumberDialog(self)
        ser_num_diag.ShowModal()

        try:
            ser_num = ser_num_diag.serial_number

            logger.info("MainWindow:new_query: get beacon info for sn# {0}".format(ser_num))
            self.statusbar.SetStatusText("Retrieving information for SN# {0}".format(ser_num))
            self.add_new_results_page(ser_num, get_beacon_info(ser_num))

        except AttributeError:
            logger.info("MainWindow:new_query: no serial number was input")

        ser_num_diag.Destroy()
        self.statusbar.SetStatusText("Ready...")

    def add_new_results_page(self, serial_number, beacon_info):
        """
            This function adds the results of a get_beacon_info() call to a page
            of the results notebook.
        :return:
        """
        logger.debug("MainWindow:add_new_results")
        self.statusbar.SetStatusText("Retrieving data from T3Production")

        if self.page_counter is 0:
            logger.debug("MainWindow:add_new_results: -> first page entry")
            self.results_notebook.DeletePage(0)
            self.results_notebook.AddPage(ResultsPage(self.results_notebook, beacon_info, serial_number), serial_number)
        else:
            logger.debug("MainWindow:add_new_results: -> add new page entry")
            self.results_notebook.AddPage(ResultsPage(self.results_notebook, beacon_info, serial_number), serial_number)

        index = self.results_notebook.GetPageCount() - 1
        self.results_notebook.SetSelection(index)
        self.page_counter += 1

        self.statusbar.SetStatusText("Done...")

    def save_results(self, e):
        """
            This function opens a save dialog window and then saves the
            selected tab to a JSON file.
        :param e:
        :return:
        """
        ser_num_to_save = self.results_notebook.GetPageText(self.results_notebook.GetSelection())
        logger.debug("MainWindow:save_results: save notebook page {0}".format(ser_num_to_save))

        if ser_num_to_save == "< ... >":
            no_report = wx.MessageDialog(None, "No report open, cannot save empty file", "Error: No report open",
                                         wx.OK | wx.ICON_ERROR)
            no_report.ShowModal()
            return

        save_diag = wx.FileDialog(self, "Save {0} file", "", "", "JSON files (*.json)|*.json", wx.FD_SAVE |
                                  wx.FD_OVERWRITE_PROMPT)

        if save_diag.ShowModal() == wx.ID_CANCEL:
            logger.debug("MainWindow:save_results: user canceled action")
            return
        else:
            logger.debug("MainWindow:save_results: path={0}".format(save_diag.GetPath()))

            # Format and save file as JSON
            save_json_file(save_diag.GetPath(), ser_num_to_save)

    def open_file(self, e):
        """
            This function opens a saved JSON file containing manufacturing
            information on a beacon and adds a tab to the main window.
        :param e:
        :return:
        """
        logger.debug("MainWindow:open_file")

        open_diag = wx.FileDialog(self, "Open JSON Report File", "", "", "JSON files (*.json)|*.json",
                                  wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if open_diag.ShowModal() == wx.ID_CANCEL:
            logger.debug("MainWindow:open_file: user canceled action")
            return
        else:
            logger.debug("Mainwindow:open_file: path={0}".format(open_diag.GetPath()))

            with open(open_diag.GetPath(), "r") as f:
                try:
                    data = json.load(f)
                except ValueError:
                    logger.debug("MainWindow:open_file: Unable to open report file")
                    return

                try:
                    # Open a new report page within the notebook
                    if data[0]["db_table"] == "assemblyKittingTable" or data[0]["db_table"] == "falloutTable":
                        serial_number_text = "serialNumberUnit"
                    else:
                        serial_number_text = "serialNumber"

                    self.add_new_results_page(data[0][serial_number_text], data)
                except IndexError:
                    logger.error("MainWindow:open_file: Unable to open file")
                    open_file_err = wx.MessageDialog(None, "Unable to open report file", "Error: Open Report File",
                                                     wx.OK | wx.ICON_ERROR)
                    open_file_err.ShowModal()

    def on_about_box(self, e):
        logger.debug("MainWindow:on_about_box")

        description = """BCA Beacon Tracker is an application for displaying manufacturing information from the
        T3 Production Database for Tracker 3 Avalanche Beacons.\r\n\nPlease report any bugs and/or feature requests
        to duncan.lowder"""

        app_license = """BCA Beacon Tracker
        Copyright (C) 2015 Backcountry Access Inc.

        This program is free software; you can redistribute it and/or
        modify it under the terms of the GNU General Public License
        as published by the Free Software Foundation; either version 2
        of the License, or (at your option) any later version.

        This program is distributed in the hope that it will be useful,
        but WITHOUT ANY WARRANTY; without even the implied warranty of
        MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
        GNU General Public License for more details.

        You should have received a copy of the GNU General Public License
        along with this program; if not, write to the Free Software
        Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA."""

        info = wx.AboutDialogInfo()

        info.SetIcon(wx.Icon("icons\\app_icon_radar.png", wx.BITMAP_TYPE_PNG))
        info.SetName("BCA Beacon Tracker")
        info.SetVersion(__version__)
        info.SetDescription(description)
        info.SetCopyright("(C) 2015 Backcountry Access Inc.")
        info.SetWebSite("http://www.backcountryaccess.com")
        info.SetLicense(app_license)
        info.AddDeveloper(__author__)

        wx.AboutBox(info)

    def on_help_box(self, e):
        """
            This function displays a dialog box containing a quick guide on
            how to use this application.
        """

        help = HelpDialog(self)

        help.ShowModal()
        help.Destroy()

    def on_quit(self, e):
        """
            This function closes and exits the application
        """
        logger.info("MainWindow:on_quit")
        self.Close()


# -----------------------------------------------------------------------------
# FUNCTIONS
# -----------------------------------------------------------------------------

def format_db_table_str(str_to_format):
    """
        This function formats the DB Table string
    :param str_to_format: string to format
    :return: formatted_str: formatted string
    """
    if str_to_format == "assemblyKittingTable":
        str_to_format = "Assembly Kitting"
    elif str_to_format == "DFTestingTable":
        str_to_format = "DF Testing"
    elif str_to_format == "calibrationTable":
        str_to_format = "Calibration"
    elif str_to_format == "finalTestTable":
        str_to_format = "Final Test"
    elif str_to_format == "closeCaseTable":
        str_to_format = "Close Case"
    elif str_to_format == "finalInspectionTable":
        str_to_format = "Final Inspection"
    elif str_to_format == "packagingTable":
        str_to_format = "Packaging"
    elif str_to_format == "falloutTable":
        str_to_format = "Fallout"
    elif str_to_format == "engineeringReworkTable":
        str_to_format = "Engineering Rework"

    return str_to_format


def format_column_str(str_to_format):
    """
        This function returns a formatted string for the different columns that
        are used within the various DB tables. This improves the human
        readability of the returned data.
    :param str_to_format: database table key string to format
    :return: formatted_string: formatted string
    """
    logger.debug("format_column_str: format {0} str".format(str_to_format))

    if str_to_format == "employeeID":
        formatted_str = "Employee"
    elif str_to_format == "workstationID":
        formatted_str = "Workstation ID"
    elif str_to_format == "failureCode":
        formatted_str = "Failure Code"
    elif str_to_format == "failureDescription":
        formatted_str = "Failure Description"
    elif str_to_format == "scanTime":
        formatted_str = "Scan Time"
    elif str_to_format == "stepResults":
        formatted_str = "Step Results"
    elif str_to_format == "LOERRORCodes":
        formatted_str = "LO Error Codes"
    elif str_to_format == "XERRORCodes":
        formatted_str = "X Error Codes"
    elif str_to_format == "YERRORCodes":
        formatted_str = "Y Error Codes"
    elif str_to_format == "serialNumber":
        formatted_str = "Serial Number"
    elif str_to_format == "serialNumberDigital":
        formatted_str = "Digital Serial Number"
    elif str_to_format == "serialNumberAnalog":
        formatted_str = "Analog Serial Number"
    elif str_to_format == "serialNumberUnit":
        formatted_str = "Top-Level Serial Number"
    elif str_to_format == "unitStatus":
        formatted_str = "Unit Status"
    else:
        formatted_str = str_to_format

    logger.debug("format_column_str: formatted str={0}".format(formatted_str))

    return formatted_str


def get_employee_name(employee_id):
    """
        This function returns a string containing the Employee Name for the ID
        that is passed in
    :param employee_id: Employee ID to return name of
    :return: workstation_str: String containing the Employees name
    """

    logger.debug("format_workstation_str: Connecting to T3Production database")
    cnxn = pyodbc.connect(sql_cnxn_str)
    cursor = cnxn.cursor()

    sql_query = "SELECT * FROM employeeTable WHERE employeeID='{0}'".format(str(employee_id))

    logger.debug("get_employee_name: SQL query={0}".format(sql_query))
    cursor.execute(sql_query)
    db_info = cursor.fetchone()

    logger.debug("get_employee_name: db_info={0}".format(db_info))

    logger.debug("get_employee_name: employee_name_str={0}".format(db_info.employeeName))

    return str(db_info.employeeName)


def get_failure_description(failure_code):
    """
        This function returns a string containing the failure description for
        the specified failure code
    :param failure_code: Code to find failure description of
    :return: failure_str: Description of the failure
    """

    logger.debug("format_workstation_str: Connecting to T3Production database")
    cnxn = pyodbc.connect(sql_cnxn_str)
    cursor = cnxn.cursor()

    sql_query = "SELECT * FROM failureModeTable WHERE failureCode='{0}'".format(str(failure_code))

    logger.debug("get_failure_description: SQL query={0}".format(sql_query))
    cursor.execute(sql_query)
    db_info = cursor.fetchone()

    logger.debug("get_failure_description: db_info={0}".format(db_info))

    logger.debug("get_failure_description: failure_str={0}".format(db_info.failureDescription))

    return str(db_info.failureDescription)


def get_db_table_info(db_table, serial_number, db_cursor):
    """
        This function returns an array containing dictionaries containing table
        entries for the specified database table and unit serial number
    :param db_table: Table to search
    :param serial_number: Serial number to search for
    :param db_cursor: Cursor for the database connection
    :return: List of dictionaries containing all of the database fields
    """
    logger.info("get_db_table_info: retrieving DB Table {0} info for serialNumber {1}".format(db_table, serial_number))

    # Set serial_number_text string, this is due to a different string being used
    # in the assemblyKittingTable and falloutTable DB Tables
    if db_table is "assemblyKittingTable" or db_table is "falloutTable":
        serial_number_text = "serialNumberUnit"
    else:
        serial_number_text = "serialNumber"

    sql_query = "SELECT * FROM {0} WHERE {1}='{2}'".format(db_table, serial_number_text, serial_number)
    logger.debug("get_db_table_info: -> execute SQL Query={0}".format(sql_query))
    db_cursor.execute(sql_query)

    # Retrieve all returned rows
    db_entries = []
    while 1:
        row = db_cursor.fetchone()
        if not row:
            break
        db_entries.append(row)
        logger.debug("get_db_table_info: -> DB Table Row={0}".format(row))

    # Get column names
    logger.info("get_db_table_info: retrieving {0} column names".format(db_table))
    db_col_names = []
    for row in db_cursor.columns(table=db_table):
        db_col_names.append(row.column_name)
    logger.debug("get_db_table_info: -> DB Table Columns={0}".format(db_col_names))

    # Generate dictionaries for the retrieved DB table entries
    logger.info("get_db_table_info: building {0} entry dictionary list".format(db_table))
    db_dict_list = []
    for entry in db_entries:

        # Add table name to dictionary
        db_dict = {"db_table": db_table}

        # Add retrieved data for DB table to dictionary
        for index in range(len(db_col_names)):
            col = db_col_names[index]
            db_dict[col] = entry[index]
        db_dict_list.append(db_dict)
        logger.debug("get_db_table_info: -> DB Dict={0}".format(db_dict))

    return db_dict_list


def get_beacon_info(serial_number):
    """
        This function returns a list containing dictionaries which contain
        information for different manufacturing steps.
    :param serial_number: Serial number of beacon to retrieve data for
    :return: list of dictionaries containing manufacturing information
    """

    # Connect to database
    logger.info("get_beacon_info: Connecting to T3Production database")
    cnxn = pyodbc.connect(sql_cnxn_str)
    cursor = cnxn.cursor()

    # Get information from the database for the specified serial number
    logger.info("get_beacon_info: Retrieving T3Production database information for {0}".format(serial_number))
    db_info = []
    for table in db_table_list:
        table_info = get_db_table_info(table, serial_number, cursor)

        # Append all entries from the DB table to the db_info list
        for index in range(len(table_info)):
            db_info.append(table_info[index])

    logger.info("get_beacon_info: Sorting retrieved data by transactionTime")
    sorted_db_info = sorted(db_info, key=itemgetter("transactionTime"))

    return sorted_db_info


def save_json_file(file_path, serial_number):
    """
        This function parses and formats the database information for the
        specified beacon and then saves it in the format of a JSON file.
    :param file_path: Location to save file
    :param serial_number: Serial number of beacon to save
    :return:
    """

    logger.debug("save_json_file: retrieving data for {0}".format(serial_number))

    data = get_beacon_info(serial_number)

    # Change datetime.datetime objects to strings
    for table_entry in data:
        table_entry['scanTime'] = str(table_entry['scanTime'])
        table_entry['transactionTime'] = str(table_entry['transactionTime'])

        logger.debug("save_json_file: scan_time={0}, transaction_time={1}".format(table_entry["scanTime"],
                                                                                  table_entry["transactionTime"]))

    with open(file_path, "w") as f:
        json.dump(data, f)


def main():
    """
        This function starts the main application and displays the wxPython
        GUI
    :return:
    """
    app = wx.App()
    MainWindow(None)
    app.MainLoop()


# -----------------------------------------------------------------------------
# RUN SCRIPT
# -----------------------------------------------------------------------------
if __name__ == '__main__':

    # Run the script
    main()