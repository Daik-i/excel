# -*- coding: utf-8 -*-
"""
http://blog.csdn.net/chenghit
"""
import wx
import os
# class MainWindow(wx.Frame):
#     """We simply derive a new class of Frame."""
#     def __init__(self, parent, title):
#         wx.Frame.__init__(self, parent, title = title, size = (200, 100))
#         self.control = wx.TextCtrl(self, style = wx.TE_MULTILINE)
#         self.CreateStatusBar()    #创建位于窗口的底部的状态栏
#
#         #设置菜单
#         filemenu = wx.Menu()
#         filemenu1 = wx.Menu()
#
#         #wx.ID_ABOUT和wx.ID_EXIT是wxWidgets提供的标准ID
#         filemenu.Append(wx.ID_ABOUT, u"关于", u"关于程序的信息")
#         filemenu.AppendSeparator()
#         filemenu.Append(wx.ID_EXIT, u"退出", u"终止应用程序")
#         menuItem = filemenu1.Append(wx.ID_SAVE, u"保存", u"保存文本")
#
#         #创建菜单栏
#         menuBar = wx.MenuBar()
#         menuBar.Append(filemenu, u"文件")
#         menuBar.Append(filemenu1, u"选项")
#
#         self.SetMenuBar(menuBar)
#         self.Show(True)
#         self.Bind(wx.EVT_MENU, self.OnAbout, menuItem)
#
#
# app = wx.App(False)
# frame = MainWindow(None, title = u"记事本")
# app.MainLoop()



# -*- coding: utf-8 -*-
"""
http://blog.csdn.net/chenghit
"""

class MainWindow(wx.Frame):
    """We simply derive a new class of Frame."""
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title = title, size = (600, 400))
        self.control = wx.TextCtrl(self, style = wx.TE_MULTILINE)
        self.CreateStatusBar()    # 创建位于窗口的底部的状态栏

        # 设置菜单
        filemenu = wx.Menu()
        filemenu1 = wx.Menu()


        # wx.ID_ABOUT和wx.ID_EXIT是wxWidgets提供的标准ID
        menuAbout = filemenu.Append(wx.ID_ABOUT, "&About", \
            " Information about this program")    # (ID, 项目名称, 状态栏信息)
        filemenu.AppendSeparator()
        menuExit = filemenu.Append(wx.ID_EXIT, "E&xit", \
            " Terminate the program")    # (ID, 项目名称, 状态栏信息)
        menuOpen = filemenu1.Append(wx.ID_OPEN, "&Open", \
                                   " Open a file")  # (ID, 项目名称, 状态栏信息)
        # 创建菜单栏
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, "&File")    # 在菜单栏中添加filemenu菜单

        menuBar.Append(filemenu1, "&Choose")
        self.SetMenuBar(menuBar)    # 在frame中添加菜单栏

        panel = wx.Panel(self)
        self.quote = wx.StaticText(panel, label="Hello!",pos=(10,10) ,size=(40,40))
        self.quote1 = wx.StaticText(panel, label="World!", pos=(10,50) ,size=(40, 40))
        # 设置events
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_MENU, self.OnOpen, menuOpen)

        # 设置sizers
        self.sizer2 = wx.BoxSizer(wx.HORIZONTAL)  #水平方向布置boxsizer
        self.buttons = []
        for i in range(0, 6):
            self.buttons.append(wx.Button(self,-1, "Button &" + str(i)))
            self.sizer2.Add(self.buttons[i], 1, wx.GROW)

        self.sizer = wx.BoxSizer(wx.VERTICAL)    #垂直方向布置boxsizer
        self.sizer.Add(panel,0, wx.GROW)
        self.sizer.Add(self.control, 1, wx.EXPAND)
        self.sizer.Add(self.sizer2, 0, wx.GROW)
        self.Bind(wx.EVT_BUTTON, self.OnOpen, self.buttons[1])

        # 激活sizer
        self.SetSizer(self.sizer)
        self.SetAutoLayout(True)
        self.sizer.Fit(self)
        self.Show(True)

    def OnAbout(self, e):
        # 创建一个带"OK"按钮的对话框。wx.OK是wxWidgets提供的标准ID
        dlg = wx.MessageDialog(self, "A small text editor.", \
            "About Sample Editor", wx.OK)    # 语法是(self, 内容, 标题, ID)
        dlg.ShowModal()    # 显示对话框
        dlg.Destroy()    # 当结束之后关闭对话框

    def OnExit(self, e):
        self.Close(True)    # 关闭整个frame

    def OnOpen(self,e):
        """ Open a file"""
        self.dirname = ''
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.*",wx.ID_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            f = open(os.path.join(self.dirname, self.filename), 'r')
            self.control.SetValue(f.read())
            f.close()
        dlg.Destroy()

app = wx.App(False)
frame = MainWindow(None, title = "Small editor")
app.MainLoop()