# -*- coding: utf-8 -*-
"""
http://blog.csdn.net/chenghit
"""

import wx


def print1():
    print('ok')
class ExamplePanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.quote = wx.StaticText(self, label='Your quote:', pos=(20, 30))
        #self.CreateStatusBar()  # 创建位于窗口的底部的状态栏
        # 这个多行的文本框只是用来记录并显示events，不要纠结之
        self.logger = wx.TextCtrl(self, pos=(300,20), size=(250,300),
                                  style=wx.TE_MULTILINE | wx.TE_READONLY)

        # 一个按钮
        self.button = wx.Button(self, label='Save', pos=(200, 325))
        self.Bind(wx.EVT_BUTTON, self.OnClick, self.button)

        # 仅有1行的编辑控件
        self.lblname = wx.StaticText(self, label='Your name:', pos=(20, 60))
        self.editname = wx.TextCtrl(self, value='Enter here your name:',
                                    pos=(100, 55), size=(140, -1))
        self.Bind(wx.EVT_TEXT, self.EvtText, self.editname)
        self.Bind(wx.EVT_CHAR, self.EvtChar, self.editname)

        # 一个ComboBox控件（下拉菜单）
        self.sampleList = ['friends', 'advertising', 'web search', \
                           'Yellow Pages']
        self.lblhear = wx.StaticText(self, label="How did you hear from us ?",
                                     pos=(20, 90))
        self.edithear = wx.ComboBox(self, pos=(200, 85), size=(95, -1),
                                    choices=self.sampleList,
                                    style=wx.CB_DROPDOWN | wx.TE_READONLY)
        self.Bind(wx.EVT_COMBOBOX, self.EvtComboBox, self.edithear)
        # 注意ComboBox也绑定了EVT_TEXT事件
        self.Bind(wx.EVT_TEXT, self.EvtText, self.edithear)

        # 复选框
        self.insure = wx.CheckBox(self, label="Do you want Insured Shipment ?",
                                  pos=(20,180))
        self.Bind(wx.EVT_CHECKBOX, self.EvtCheckBox, self.insure)

        # 单选框
        radioList = ['blue', 'red', 'yellow', 'orange', 'green', 'purple', \
                     'navy blue', 'black', 'gray']
        self.rb = wx.RadioBox(self,label="What color would you like ?",\
                              pos=(20, 210), choices=radioList, \
                              majorDimension=3, style=wx.RA_SPECIFY_COLS)
        self.Bind(wx.EVT_RADIOBOX, self.EvtRadioBox, self.rb)

    def OnClick(self, event):
        self.logger.AppendText('Click on object with Id %d\n' % \
                               event.GetId())
    def EvtText(self, event):
        self.logger.AppendText('EvtText: %s\n' % event.GetString())
    def EvtChar(self, event):
        self.logger.AppendText('EvtChar: %d\n' % event.GetKeyCode())
        event.Skip()
    def EvtComboBox(self, event):
        self.logger.AppendText('EvtComboBox: %s\n' % event.GetString())
    def EvtCheckBox(self, event):
        self.logger.AppendText('EvtCheckBox: %d\n' % event.IsChecked())
    def EvtRadioBox(self, event):
        self.logger.AppendText('EvtRadioBox: %d\n' % event.GetInt())
        print1()





app = wx.App(False)
frame = wx.Frame(None, size=(600,400))
panel = ExamplePanel(frame)
frame.Show()
app.MainLoop()