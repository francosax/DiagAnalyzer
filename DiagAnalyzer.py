import wx
import os
from CoreDiagAnalyzer_v1 import *

aboutText = """<p>Sorry, there is no information about this program. It is
running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>"""

# class HtmlWindow(wx.html.HtmlWindow):
#     def __init__(self, parent, id, size=(600,400)):
#         wx.html.HtmlWindow.__init__(self,parent, id, size=size)
#         if "gtk2" in wx.PlatformInfo:
#             self.SetStandardFonts()
#
#     def OnLinkClicked(self, link):
#         wx.LaunchDefaultBrowser(link.GetHref())
#
# class AboutBox(wx.Dialog):
#     def __init__(self):
#         wx.Dialog.__init__(self, None, -1, "About <<project>>",
#             style=wx.DEFAULT_DIALOG_STYLE|wx.THICK_FRAME|wx.RESIZE_BORDER|
#                 wx.TAB_TRAVERSAL)
#         hwin = HtmlWindow(self, -1, size=(400,200))
#         vers = {}
#         vers["python"] = sys.version.split()[0]
#         vers["wxpy"] = wx.VERSION_STRING
#         hwin.SetPage(aboutText % vers)
#         btn = hwin.FindWindowById(wx.ID_OK)
#         irep = hwin.GetInternalRepresentation()
#         hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
#         self.SetClientSize(hwin.GetSize())
#         self.CentreOnParent(wx.BOTH)
#         self.SetFocus()

class MyForm(wx.Frame):
    """Constructor"""
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, title='Diag Analyzer v1', pos=(150, 150), size=(700, 320))
        self.Bind(wx.EVT_CLOSE, self.OnCloseFrame)

        # Add a panel so it looks correct on all platforms
        # panel = wx.Panel(self)
        self.panel = wx.Panel(self, wx.ID_ANY)  # , pos=(150, 150), size=(675, 300))
        bmp = wx.ArtProvider.GetBitmap(wx.ART_FILE_OPEN, wx.ART_OTHER, (16, 16))
        title_ico = wx.StaticBitmap(self.panel, wx.ID_ANY, bmp)
        title = wx.StaticText(self.panel, wx.ID_ANY, 'Req Diagnosis Analyzer')
        bmp = wx.ArtProvider.GetBitmap(wx.ART_TIP, wx.ART_OTHER, (16, 16))

        # Add a Status Bar
        self.status_string = ""
        self.stat_bar = self.CreateStatusBar(2)
        self.stat_bar.SetStatusWidths([-1, 120])
        # self.stat_bar = wx.StatusBar(self.panel, wx.ID_ANY, wx.STB_DEFAULT_STYLE, name="Operation")

        # Add a input box
        input_one_ico = wx.StaticBitmap(self.panel, wx.ID_ANY, bmp)
        label_one = wx.StaticText(self.panel, wx.ID_ANY, 'Directory')
        self.inputTxtOne = wx.TextCtrl(self.panel, wx.ID_ANY, 'D:\\users\\fsacchet\\PycharmProjects\\MyDataAna')

        # Add a input box
        input_two_ico = wx.StaticBitmap(self.panel, wx.ID_ANY, bmp)
        label_two = wx.StaticText(self.panel, wx.ID_ANY, 'first file')
        self.inputTxtTwo = wx.TextCtrl(self.panel, wx.ID_ANY, 'Req_SYSDIAG_P230_EMEA_20062019.xlsx')

        # Add a input box
        input_three_ico = wx.StaticBitmap(self.panel, wx.ID_ANY, bmp)
        label_three = wx.StaticText(self.panel, wx.ID_ANY, 'second file')
        self.inputTxtThree = wx.TextCtrl(self.panel, wx.ID_ANY, 'Req_SYSDIAG_P230_NAFTA_20062019.xlsx')

        # Add a OK button
        self.okBtn = wx.Button(self.panel, wx.ID_ANY, 'OK')
        # Add a CANCEL button
        self.cancelBtn = wx.Button(self.panel, wx.ID_ANY, 'Cancel')
        # Bind events to the buttons
        self.Bind(wx.EVT_BUTTON, self.onOK, self.okBtn)
        self.Bind(wx.EVT_BUTTON, self.onCancel, self.cancelBtn)

        # try to have a help menu and Exit menu
        self.menuBar = wx.MenuBar()

        # File menu
        file_menu = wx.Menu()
        exit_item = file_menu.Append(wx.ID_EXIT, "E&xit\tCtrl-Q", "Exit Diagnosis Analyzer")
        # exit_item.SetBitmap(images.Exit.GetBitmap())
        self.Bind(wx.EVT_MENU, self.OnQuit, exit_item)
        self.menuBar.Append(file_menu, "&File")

        # Help/About menu
        help_menu = wx.Menu()
        about_item = help_menu.Append(wx.ID_ABOUT, '&About Diagnosis Analyzer', 'About')
        help_item = help_menu.Append(wx.ID_HELP, '&Help', 'Help')
        self.Bind(wx.EVT_MENU, self._on_help_about, about_item)
        self.Bind(wx.EVT_MENU, self._on_help_help, help_item)
        self.menuBar.Append(help_menu, '&Help')
        self.SetMenuBar(self.menuBar)

        # Sizers allocations
        top_sizer = wx.BoxSizer(wx.VERTICAL)
        title_sizer = wx.BoxSizer(wx.HORIZONTAL)
        input_one_sizer = wx.BoxSizer(wx.HORIZONTAL)
        input_two_sizer = wx.BoxSizer(wx.HORIZONTAL)
        input_three_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        status_bar_sizer = wx.BoxSizer(wx.HORIZONTAL)

        title_sizer.Add(title_ico, 0, wx.ALL, 5)
        title_sizer.Add(title, 0, wx.ALL, 5)

        input_one_sizer.Add(input_one_ico, 0, wx.ALL, 5)
        input_one_sizer.Add(label_one, 0, wx.ALL, 5)
        input_one_sizer.Add(self.inputTxtOne, 1, wx.ALL | wx.EXPAND, 5)

        input_two_sizer.Add(input_two_ico, 0, wx.ALL, 5)
        input_two_sizer.Add(label_two, 0, wx.ALL, 5)
        input_two_sizer.Add(self.inputTxtTwo, 1, wx.ALL | wx.EXPAND, 5)

        input_three_sizer.Add(input_three_ico, 0, wx.ALL, 5)
        input_three_sizer.Add(label_three, 0, wx.ALL, 5)
        input_three_sizer.Add(self.inputTxtThree, 1, wx.ALL | wx.EXPAND, 5)

        btn_sizer.Add(self.okBtn, 0, wx.ALIGN_BOTTOM | wx.EXPAND, 5)
        btn_sizer.Add(self.cancelBtn, 0, wx.ALIGN_BOTTOM, 5)

        # status bar
        # status_bar_sizer.Add(self.stat_bar, 0, wx.ALL, 5)

        top_sizer.Add(title_sizer, 0, wx.CENTER)
        top_sizer.Add(wx.StaticLine(self.panel,), 0, wx.ALL | wx.EXPAND, 5)
        top_sizer.Add(input_one_sizer, 0, wx.ALL | wx.EXPAND, 5)
        top_sizer.Add(input_two_sizer, 0, wx.ALL | wx.EXPAND, 5)
        top_sizer.Add(input_three_sizer, 0, wx.ALL | wx.EXPAND, 5)
        top_sizer.Add(wx.StaticLine(self.panel), 0, wx.ALL | wx.EXPAND, 5)
        top_sizer.Add(btn_sizer, 0, wx.ALL | wx.CENTER | wx.ALIGN_BOTTOM, 5)
        top_sizer.Add(status_bar_sizer, 0, wx.ALL | wx.EXPAND, 5)

        self.panel.SetSizer(top_sizer)
        self.panel.Layout()
        # top_sizer.Fit(self)

        window = top_sizer.ComputeFittingWindowSize(self)  # compute window actual size
        print("top_sizer best calculated size:", window)

        # win_obj.SetDimension(300, 500, 300, 500)
        # top_sizer.Fit(self)
        # top_sizer.SetDimension(300, 500, 300, 500)

        # top_sizer.RecalcSizes() SetMinSize()

    def set_temp_status(self, event):
        self.stat_bar.PushStatusText(self.status_string)

    def restore_status(self, event):
        self.stat_bar.PopStatusText()

    def onOK(self, event):
        # Do something
        op_status = ''
        try:
            self.second_file = self.inputTxtThree.GetValue()
            self.path = self.inputTxtOne.GetValue()
            self.first_file = self.inputTxtTwo.GetValue()
            # Bind events to the status bar start
            self.stat_bar.SetStatusText("Working in: " + self.path, 0)
            # Bind events to the status bar end
            os.chdir(self.path)
            self.stat_bar.SetStatusText("Operation in progress", 1)
            [df1, df2, df3, df4, difference_dtc, difference_rec, op_status] = \
                write_excel_df_diff(self.first_file, self.second_file)
            self.stat_bar.SetStatusText("DONE", 1)
            if 'OK' in op_status:
                self.stat_bar.SetStatusText("Result: " + op_status, 0)
                self.stat_bar.SetStatusText("DONE", 1)
            else:
                self.stat_bar.SetStatusText("Result: " + op_status + ', please check file names and directory', 0)
                self.stat_bar.SetStatusText("DONE", 1)
        except IOError:
            # Bind events to the status bar start
            self.status_string = "Wrong operation"
            self.set_temp_status(self.status_string)
            self.stat_bar.SetStatusText('Open error: please check directory name ' + op_status, 0)
            self.stat_bar.SetStatusText(self.status_string, 1)
            # Bind events to the status bar end
        print('onOK handler done')

    def onCancel(self, event):
        self.inputTxtOne.SetValue('')
        self.inputTxtTwo.SetValue('')
        self.inputTxtThree.SetValue('')
        # self.closeProgram()

    def closeProgram(self):
        """Destructor"""
        self.Destroy()
        print("Destroy")
        self.DeletePendingEvents()
        print("Deleted Pending Events")
        self.Close()
        print("Close")

    # Bind our events from our menu item and from the close dialog 'x' on the frame
    def SetupEvents(self):
        self.Bind(wx.EVT_CLOSE, self.OnCloseFrame)
        self.Bind(wx.EVT_MENU, self.OnCloseFrame, self.fileMenuExitItem)

    # Destroys the main frame which quits the wxPython application
    def OnExitApp(self, event):
        self.Destroy()

    # Makes sure the user was intending to quit the application
    def OnCloseFrame(self, event):
        dialog = wx.MessageDialog(self, message="Are you sure you want to quit?", caption="Exit Diagnosis Analyzer",
                                  style=wx.YES_NO, pos=wx.DefaultPosition)
        response = dialog.ShowModal()

        if response == wx.ID_YES:
            self.OnExitApp(event)
        else:
            event.StopPropagation()

    def OnQuit(self, event):
        self.Close()

    def _on_help_about(self, event):
        print("this is an about")

    def _on_help_help(self, event):
        print("this is an help")


# ------------------------------------------------
#   TEST MAIN
# ------------------------------------------------
#
# Run the program
if __name__ == '__main__':
    import wx.lib.inspection
    app = wx.App()
    mf = MyForm()
    # mf.SetAutoLayout(1) # autorescale all off
    try:
        mf.Show()
        # mf.register_close_callback(lambda: True)
        # wx.lib.inspection.InspectionTool().Show()
        app.MainLoop()
    except PermissionError:
        mf.DeletePendingEvents()
        mf.Destroy()
