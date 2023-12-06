#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project:Factory_test
@File:serial_handler.py
@Author:rivern.yuan
@Date:2022/9/15 21:09
"""

import sys
import os
import time
import wx
import locale
import gettext
import asyncio
import datetime
import subprocess
import threading
import file_handler
import serial_list
import serial_handler
import configparser
from wx import stc
from log_redirect import RedirectErr, RedirectStd
from pubsub import pub
from queue import Queue
import codecs
import json

if hasattr(sys, '_MEIPASS'):
    PROJECT_ABSOLUTE_PATH = os.path.dirname(os.path.realpath(sys.executable))
else:
    PROJECT_ABSOLUTE_PATH = os.path.dirname(os.path.abspath(__file__))


def set_config(PROJECT_ABSOLUTE_PATH, frame):
    if not os.path.exists(PROJECT_ABSOLUTE_PATH+"\\config.ini"):
        return
    conf = configparser.ConfigParser(interpolation=None)
    conf.read(PROJECT_ABSOLUTE_PATH+"\\config.ini", encoding='utf-8')
    py_file = conf.get("software", "py_file")
    json_file = conf.get("software", "json_file")

    if py_file or json_file:
        no_exist = []
        msg = wx.MessageBox(_(u'是否恢复上次文件'), _(u'提示'), wx.YES_NO)
        if msg == wx.YES:
            if py_file:
                if os.path.exists(py_file):
                    try:
                        relative_path = os.path.relpath(py_file, PROJECT_ABSOLUTE_PATH)
                        frame.text_ctrl_py.SetValue(relative_path)
                    except Exception as e:
                        frame.text_ctrl_py.SetValue(py_file)
                else:
                    conf.set("software", "py_file", "")
                    wx.MessageBox(py_file + _(u"文件不存在"), _(u"提示"), wx.YES_DEFAULT | wx.ICON_INFORMATION)

            if json_file:
                if os.path.exists(json_file):
                    try:
                        relative_path = os.path.relpath(json_file, PROJECT_ABSOLUTE_PATH)
                        frame.text_ctrl_json.SetValue(relative_path)
                    except Exception as e:
                        frame.text_ctrl_json.SetValue(json_file)
                    frame.load_json_file(0, json_file)
                else:
                    conf.set("software", "json_file", "")
                    wx.MessageBox(json_file + _(u"文件不存在"), _(u"提示"), wx.YES_DEFAULT | wx.ICON_INFORMATION)

            with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
                conf.write(f)

        elif msg == wx.NO:
            conf.set("software", "py_file", "")
            conf.set("software", "json_file", "")
            with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
                conf.write(f)


class AdvancedSettings(wx.Dialog):
    # automated, self.testMessages, selectedList
    def __init__(self, parent, automated, selectedList, json_file):
        super(AdvancedSettings, self).__init__(parent, title="Advanced Settings", size=(350, 200))
        self.SetSize(wx.DLG_UNIT(self, wx.Size(400, 280)))

        if json_file:
            with codecs.open(json_file, 'r', 'utf-8') as f:
                data = json.load(f)
                info = data["info"]
                self.funcList = [i[0] for i in info]
        else:
            self.funcList = []

        # self.funcList = funcList
        self.parent = parent
        self.add_partition_box = wx.CheckBox(self, 100, _(u"自动开始测试"))
        if automated == "True":
            self.add_partition_box.SetValue(True)
        else:
            self.add_partition_box.SetValue(False)

        self.restart_box = wx.CheckBox(self, 107, _(u"自动重启"))
        if restartSign == "True":
            self.restart_box.SetValue(True)
        else:
            self.restart_box.SetValue(False)

        self.imei_box = wx.CheckBox(self, wx.ID_ANY, _(u"Imei号匹配"))
        if self.parent.conf.get("detection", "imei") == "True":
            self.imei_box.SetValue(True)
        else:
            self.imei_box.SetValue(False)

        self.version_box = wx.CheckBox(self, 106, _(u"版本号检测"))
        self.label_1 = wx.StaticText(self, wx.ID_ANY, _(u"cmd: "))
        self.cmd_txt = wx.TextCtrl(self, wx.ID_ANY, self.parent.conf.get("version", "cmd"))
        self.label_2 = wx.StaticText(self, wx.ID_ANY, _(u"版本号: "))
        self.version_txt = wx.TextCtrl(self, wx.ID_ANY, self.parent.conf.get("version", "version"))
        if versionDetection == "True":
            self.version_box.SetValue(True)
        else:
            self.version_box.SetValue(False)

        self.funcSelector = wx.Choice(self, wx.ID_ANY, choices=self.funcList)
        self.selectedListBox = wx.ListBox(self, wx.ID_ANY, choices=selectedList)
        self.add = wx.Button(self, 101, _(u"添加"))
        self.delete = wx.Button(self, 102, _(u"删除"))
        self.clear = wx.Button(self, 103, _(u"清空"))
        self.selectAll = wx.Button(self, 104, _(u"全选"))
        self.confirm = wx.Button(self, 105, _(u"确认"))
        self.buttonFotaCancel = wx.Button(self, wx.ID_CANCEL, _(u"取消"))

        # self.Bind(wx.EVT_CHECKBOX, self.key_processing, self.add_partition_box)
        # self.Bind(wx.EVT_CHECKBOX, self.key_processing, self.version_box)
        self.Bind(wx.EVT_BUTTON, self.key_processing, self.add)
        self.Bind(wx.EVT_BUTTON, self.key_processing, self.delete)
        self.Bind(wx.EVT_BUTTON, self.key_processing, self.clear)
        self.Bind(wx.EVT_BUTTON, self.key_processing, self.selectAll)
        self.Bind(wx.EVT_BUTTON, self.key_processing, self.confirm)

        self.__set_properties()
        self.__do_layout()
        self.Center()

    def __set_properties(self):
        self.funcSelector.SetMinSize((150, 25))
        self.selectedListBox.SetMinSize((150, 300))

        # self.label_1.SetMinSize((50, 25))
        self.cmd_txt.SetMinSize((200, 20))
        # self.label_1.SetMinSize((50, 25))
        self.version_txt.SetMinSize((200, 20))

    def __do_layout(self):
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_3_ = wx.BoxSizer(wx.HORIZONTAL)
        sizer_3_1 = wx.BoxSizer(wx.HORIZONTAL)


        sizer_3.Add((20, 0), 0, 0, 0)
        sizer_3.Add(self.add_partition_box, 0, 0, 0)

        sizer_3_.Add((20, 0), 0, 0, 0)
        sizer_3_.Add(self.restart_box, 0, 0, 0)

        sizer_3_1.Add((20, 0), 0, 0, 0)
        sizer_3_1.Add(self.imei_box, 0, 0, 0)

        sizer_4.Add((20, 0), 0, 0, 0)
        sizer_4.Add(self.version_box, 0, 0, 0)
        sizer_4.Add((20, 0), 0, 0, 0)
        sizer_4.Add(self.label_1, 0, 0, 0)
        sizer_4.Add(self.cmd_txt, 0, 0, 0)
        sizer_4.Add((20, 0), 0, 0, 0)
        sizer_4.Add(self.label_2, 0, 0, 0)
        sizer_4.Add(self.version_txt, 0, 0, 0)

        sizer_box_1 = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, _(u" 测项配置 ")), wx.HORIZONTAL)


        sizer_2.Add(self.add, 0, 0, 0)
        sizer_2.Add((0, 20), 0, 0, 0)
        sizer_2.Add(self.delete, 0, 0, 0)
        sizer_2.Add((0, 20), 0, 0, 0)
        sizer_2.Add(self.clear, 0, 0, 0)
        sizer_2.Add((0, 20), 0, 0, 0)
        sizer_2.Add(self.selectAll, 0, 0, 0)
        # sizer_2.Add((0, 20), 0, 0, 0)
        # sizer_2.Add(self.confirm, 0, 0, 0)

        sizer_box_1.Add(self.funcSelector, 0, 0, 0)
        sizer_box_1.Add((20, 0), 0, 0, 0)
        sizer_box_1.Add(sizer_2, 0, 0, 0)
        sizer_box_1.Add((20, 0), 0, 0, 0)
        sizer_box_1.Add(self.selectedListBox, 0, 0, 0)

        sizer_1.Add((0, 10), 0, 0, 0)
        sizer_1.Add(sizer_3, 0, 0, 0)
        sizer_1.Add((0, 20), 0, 0, 0)
        sizer_1.Add(sizer_3_, 0, 0, 0)
        sizer_1.Add((0, 20), 0, 0, 0)
        sizer_1.Add(sizer_3_1, 0, 0, 0)
        sizer_1.Add((0, 20), 0, 0, 0)
        sizer_1.Add(sizer_4, 0, 0, 0)
        sizer_1.Add((0, 20), 0, 0, 0)
        sizer_1.Add(sizer_box_1, 0, 0, 0)

        sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_5.Add(self.confirm, 0, 0, 0)
        sizer_5.Add(self.buttonFotaCancel, 0, wx.LEFT, 10)
        sizer_1.Add((0, 10), 0, 0, 0)
        sizer_1.Add(sizer_5, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        self.SetSizer(sizer_1)

    def key_processing(self, event):

        if event.GetId() == 101:
            funcName = self.funcSelector.GetString(self.funcSelector.GetSelection())

            if funcName and self.selectedListBox.FindString(funcName, caseSensitive = True) == -1:
                self.selectedListBox.Append(funcName)

        elif event.GetId() == 102:
            deleted_item = self.selectedListBox.GetSelection()
            if deleted_item != -1:
                projectName = self.selectedListBox.GetString(deleted_item)
                print(projectName)
                self.selectedListBox.Delete(deleted_item)

        elif event.GetId() == 103:
            self.selectedListBox.Clear()

        elif event.GetId() == 104:
            self.selectedListBox.Clear()
            for x in self.funcList:
                self.selectedListBox.Append(x)

        elif event.GetId() == 105:
            global selectedList, automated, versionDetection, restartSign

            # 自动开始测试配置
            if self.add_partition_box.GetValue():
                automated = "True"
            else:
                automated = "False"

            self.parent.conf.set("detection", "automated", automated)

            # 自动重启配置
            if self.restart_box.GetValue():
                restartSign = "True"
            else:
                restartSign = "False"

            self.parent.conf.set("detection", "restart", restartSign)

            # Imei号匹配配置
            if self.imei_box.GetValue():
                imei = "True"
            else:
                imei = "False"

            self.parent.conf.set("detection", "imei", imei)

            # 版本号测试配置
            if self.version_box.GetValue():
                versionDetection = "True"
            else:
                versionDetection = "False"
            self.parent.conf.set("version", "ifdetect", versionDetection)
            self.parent.conf.set("version", "cmd", self.cmd_txt.GetValue())
            self.parent.conf.set("version", "version", self.version_txt.GetValue())

            # 测项配置
            selectedList = []
            for i in range(self.selectedListBox.GetCount()):
                projectName = self.selectedListBox.GetString(i)
                selectedList.append(projectName)

            self.parent.conf.set("detection", "item", ",".join(selectedList))
            with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
                self.parent.conf.write(f)

            json_file = self.parent.text_ctrl_json.GetValue()
            if json_file:
                self.parent.load_json_file(0, json_file)
            self.Destroy()


class FactoryFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        kwargs["style"] = kwargs.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwargs)
        self.SetTitle("FactoryTool")
        _icon = wx.NullIcon
        _icon.CopyFromBitmap(wx.Bitmap(PROJECT_ABSOLUTE_PATH + "\\media\\quectel.ico", wx.BITMAP_TYPE_ICO))
        self.SetIcon(_icon)
        self.SetSize((1120, 700))
        self.Center()
        # stdout log
        sys.stderr = RedirectErr(self, PROJECT_ABSOLUTE_PATH)
        sys.stdout = RedirectStd(self, PROJECT_ABSOLUTE_PATH)
        # module log
        self.logger = ModuleLog()
        # Menu Bar
        self.menuBar = wx.MenuBar()
        wxg_tmp_menu = wx.Menu()
        # wxg_tmp_menu.Append(1001, u"Auto detect Port", "")
        # self.Bind(wx.EVT_MENU, self.__menu_handler, id=1001)
        # wxg_tmp_menu.AppendSeparator()
        wxg_tmp_menu.Append(1002, _(u"编辑Py文件"), "")
        self.Bind(wx.EVT_MENU, self.__menu_handler, id=1002)
        wxg_tmp_menu.Append(1003, _(u"编辑Excel文件"), "")
        self.Bind(wx.EVT_MENU, self.__menu_handler, id=1003)
        self.menuBar.Append(wxg_tmp_menu, _(u"编辑"))
        wxg_tmp_menu_1 = wx.Menu()
        wxg_tmp_menu_1.Append(2001, _(u"工具日志"), "")
        self.Bind(wx.EVT_MENU, self.__menu_handler, id=2001)
        wxg_tmp_menu_1.AppendSeparator()
        wxg_tmp_menu_1.Append(2002, _(u"模块日志"), "")
        self.Bind(wx.EVT_MENU, self.__menu_handler, id=2002)
        self.menuBar.Append(wxg_tmp_menu_1, _(u"日志"))


        wxg_tmp_menu_2 = wx.Menu()
        wxg_tmp_menu_2.Append(3001, _(u"高级设置"), "")
        self.Bind(wx.EVT_MENU, self.__menu_handler, id=3001)
        self.menuBar.Append(wxg_tmp_menu_2, _(u"设置"))

        self.SetMenuBar(self.menuBar)
        # Menu Bar end
        # self.statusBar = self.CreateStatusBar(3)
        self.main_panel = wx.Panel(self, wx.ID_ANY)
        self.load_py = wx.Button(self.main_panel, wx.ID_ANY, _("选择"))
        self.text_ctrl_py = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.load_json = wx.Button(self.main_panel, wx.ID_ANY, _("选择"))
        self.text_ctrl_json = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)

        self.button_all_start = wx.Button(self.main_panel, wx.ID_ANY, _("全部开始"))

        self.button_start_list = []
        self.ListCtrl_list = []
        self.port_ctrl_list = []
        self.Imei_ctrl_list = []
        self.tips_ctrl_list = []
        for i in range(4):
            ListCtrl = wx.ListCtrl(self.main_panel, 200+i, style=wx.LC_HRULES | wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_VRULES)
            ListCtrl.SetMinSize((-1, 250))
            ListCtrl.AppendColumn(_(u"测项"), format=wx.LIST_FORMAT_LEFT, width=130)
            ListCtrl.AppendColumn(_(u"测试方式"), format=wx.LIST_FORMAT_LEFT, width=70)
            ListCtrl.AppendColumn(_(u"测试结果"), format=wx.LIST_FORMAT_LEFT, width=70)

            tips_ctrl = wx.TextCtrl(self.main_panel, wx.ID_ANY, style=wx.TE_MULTILINE|wx.TE_WORDWRAP|wx.TE_READONLY|wx.TE_RICH2)
            button = wx.Button(self.main_panel, 100+i, _("开始"))
            button.Enable(False)
            text_ctrl = wx.TextCtrl(self.main_panel, 300+i, "", style=wx.TE_READONLY)
            text_ctrl_imei = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)

            # tips_ctrl.AppendText("请按键")

            self.ListCtrl_list.append(ListCtrl)
            self.button_start_list.append(button)
            self.port_ctrl_list.append(text_ctrl)
            self.Imei_ctrl_list.append(text_ctrl_imei)
            self.tips_ctrl_list.append(tips_ctrl)

        self.__exec_py_cmd_list = ["from usr.test import TestBase"]
        self.__init_excel()

        # event message queue
        self.message_queue = Queue(maxsize=10)
        self._channel_list = []
        self.testMessages = []
        threading.Thread(target=self.port_process_handler, args=()).start()

        # bind Func
        self.Bind(wx.EVT_BUTTON, self.__load_py_file, self.load_py)
        self.Bind(wx.EVT_BUTTON, self.load_json_file, self.load_json)
        self.Bind(wx.EVT_BUTTON, self.test_all_start, self.button_all_start)


        self.procese_num = 0

        for i in self.button_start_list:
            self.Bind(wx.EVT_BUTTON, self.test_start, i)
        self.Bind(wx.EVT_CLOSE, self.close_window)

        # pub message
        pub.subscribe(self.port_update, "serialUpdate")

        self.conf = configparser.ConfigParser(interpolation=None)
        if not os.path.exists(PROJECT_ABSOLUTE_PATH+"\\config.ini"):
            self.initConfigFile()
        self.conf.read(PROJECT_ABSOLUTE_PATH+"\\config.ini", encoding='utf-8')

        global automated, selectedList, versionDetection, restartSign
        automated = self.conf.get("detection", "automated")
        restartSign = self.conf.get("detection", "restart")
        selectedList = [i for i in self.conf.get("detection", "item").split(",") if i != ""]
        versionDetection = self.conf.get("version", "ifdetect")

        self.__set_properties()
        self.__do_layout()

    def __set_properties(self):
        for i in range(4):
            self.button_start_list[i].SetMinSize((150, 50))
            self.port_ctrl_list[i].SetMinSize((80, 25))
            self.Imei_ctrl_list[i].SetMinSize((190, 25))
            self.tips_ctrl_list[i].SetMinSize((250, 100))

        self.button_all_start.SetMinSize((150, 50))

    def __do_layout(self):
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_3 = wx.StaticBoxSizer(wx.StaticBox(self.main_panel, wx.ID_ANY, _("选择测试脚本")), wx.HORIZONTAL)
        sizer_4 = wx.StaticBoxSizer(wx.StaticBox(self.main_panel, wx.ID_ANY, _("选择配置文件")), wx.HORIZONTAL)
        sizer_5 = wx.StaticBoxSizer(wx.StaticBox(self.main_panel, wx.ID_ANY, _("工厂测试")), wx.HORIZONTAL)

        sizer_3.Add((4, 20), 0, 0, 0)
        sizer_3.Add(self.load_py, 0, 0, 0)
        sizer_3.Add((20, 20), 0, 0, 0)
        sizer_3.Add(self.text_ctrl_py, 1, 0, 0)
        sizer_3.Add((20, 20), 0, 0, 0)
        # sizer_2.Add((5, 20), 0, 0, 0)
        sizer_2.Add(sizer_3, 0, wx.EXPAND, 0)

        sizer_4.Add((4, 20), 0, 0, 0)
        sizer_4.Add(self.load_json, 0, 0, 0)
        sizer_4.Add((20, 20), 0, 0, 0)
        sizer_4.Add(self.text_ctrl_json, 1, 0, 0)
        sizer_4.Add((20, 20), 0, 0, 0)
        sizer_2.Add((20, 4), 0, 0, 0)
        sizer_2.Add(sizer_4, 0, wx.EXPAND, 0)

        for i in range(4):
            sizer_6 = wx.BoxSizer(wx.VERTICAL)
            sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_7.Add(self.port_ctrl_list[i], 0, 0, 1)
            sizer_7.Add(self.Imei_ctrl_list[i], 0, 0, 1)
            sizer_6.Add(sizer_7, 0, 0, 1)
            sizer_6.Add(self.ListCtrl_list[i], 1, 1, 1)
            sizer_6.Add((0, 10), 0, 0, 0)
            sizer_6.Add(self.tips_ctrl_list[i], 0, 1, 0)
            sizer_6.Add((0, 10), 0, 0, 0)
            sizer_6.Add(self.button_start_list[i], 0, 1, 0)
            sizer_5.Add(sizer_6, 1, 0, 0)
        sizer_2.Add(sizer_5, 0, wx.EXPAND, 0)

        sizer_2.Add((0, 10), 0, 1, 1)
        sizer_2.Add(self.button_all_start, 0, 1, 1)

        self.main_panel.SetSizer(sizer_2)
        sizer_1.Add(self.main_panel, 1, wx.EXPAND | wx.TOP, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # sizer_5.SetLabel(_("测试中"))

    def __menu_handler(self, event):  # wxGlade: FactoryFrame.<event_handler>
        if event.GetId() == 1001:
            DialogControl(self, "demo").Show()

        elif event.GetId() == 1002:
            file_path, app_path = "notepad.exe", PROJECT_ABSOLUTE_PATH + "\\module_test.py"
            raw_file_path = r'%s' % file_path
            raw_app_path = r'%s' % app_path
            subprocess.call([raw_file_path, raw_app_path])
        elif event.GetId() == 1003:
            excel_file = subprocess.Popen(["start", "/WAIT", PROJECT_ABSOLUTE_PATH + "\\Test-Result.xlsx"], shell=True)
            # psutil.Process(excel_file.pid).get_children()[0].kill()
            excel_file.poll()
        elif event.GetId() == 2001:
            p = subprocess.Popen("explorer.exe " + PROJECT_ABSOLUTE_PATH + "\\logs\\software\\", shell=True)
        elif event.GetId() == 2002:
            p = subprocess.Popen("explorer.exe " + PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\", shell=True)
        elif event.GetId() == 3001:
            global automated, selectedList
            AdvancedSettings(self, automated, selectedList, self.text_ctrl_json.GetValue()).ShowModal()
        event.Skip()

    def initConfigFile(self):
        self.conf.add_section("software")
        self.conf.set("software", "py_file", "")
        self.conf.set("software", "json_file", "")
        self.conf.add_section("detection")
        self.conf.set("detection", "automated", "False")
        self.conf.set("detection", "item", "")
        self.conf.set("detection", "restart", "False")
        self.conf.set("detection", "imei", "False")
        self.conf.add_section("version")
        self.conf.set("version", "ifdetect", "False")
        self.conf.set("version", "cmd", "import uos\\r\\nuos.uname()")
        self.conf.set("version", "version", "")
        with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
            self.conf.write(f)

    def __load_py_file(self, event):
        defDir, defFile = '', ''  # default dir/ default file
        dlg = wx.FileDialog(self, u'Open Py File', defDir, defFile, 'Python file (*.py)|*.py',
                            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return
        py_cmd = file_handler.ScriptHandler(dlg.GetPath())
        if py_cmd.script_check():
            try:
                relative_path = os.path.relpath(dlg.GetPath(), PROJECT_ABSOLUTE_PATH)
                self.text_ctrl_py.SetValue(relative_path)
            except Exception as e:
                self.text_ctrl_py.SetValue(dlg.GetPath())

            self.conf.set("software", "py_file", dlg.GetPath())
            with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
                self.conf.write(f)
        else:
            wx.MessageBox(_(u'py代码有语法错误'), u'Warn', wx.YES_DEFAULT | wx.ICON_INFORMATION)

    def load_json_file(self, event, json_file=None):
        print("json button is press")
        if event != 0:
            defDir, defFile = '', ''  # default dir/ default file
            dlg = wx.FileDialog(self, u'Open Config File', defDir, defFile, 'Config file (*.json)|*.json',
                                wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
            if dlg.ShowModal() == wx.ID_CANCEL:
                return
            json_file = dlg.GetPath()
        try:
            with codecs.open(json_file, 'r', 'utf-8') as f:
                data = json.load(f)
                print(data)
                info = data["info"]
                if selectedList:
                    info = [i for i in info if i[0] in selectedList]
                self.testFunctions = [i[1:] for i in info]
                self.testMessages = [i[0] for i in info]
                self.tips_info = {}
                try:
                    self.tips_info = data["tips"]
                except Exception as e:
                    print("523 e: ", e)

                print("versionDetection: ", versionDetection)
                if versionDetection == "True":
                    self.testFunctions.append([self.conf.get("version", "cmd"), 0, 0])
                    self.testMessages.append(_("版本号检测"))

                if self.conf.get("detection", "imei") == "True":
                    self.testFunctions.append(["Imei号匹配", 0, 0])
                    self.testMessages.append(_("Imei号匹配"))

            try:
                relative_path = os.path.relpath(json_file, PROJECT_ABSOLUTE_PATH)
                self.text_ctrl_json.SetValue(relative_path)
            except Exception as e:
                self.text_ctrl_json.SetValue(json_file)

            self.conf.set("software", "json_file", json_file)
            with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
                self.conf.write(f)

            for i, port_ctrl in enumerate(self.port_ctrl_list):
                if port_ctrl.GetValue():
                    self.ListCtrl_list[i].DeleteAllItems()
                    self.Imei_ctrl_list[i].SetValue("")
                    self.tips_ctrl_list[i].SetValue("")
                    for j, element in enumerate(self.testFunctions):
                        ListCtrl = self.ListCtrl_list[i]
                        # testFunction = self.testFunctions[j]
                        ListCtrl.InsertItem(j, j)
                        ListCtrl.SetItem(j, 0, self.testMessages[j])
                        if element[-1] == 1:
                            ListCtrl.SetItem(j, 1, _(u"人工检测"))
                        else:
                            ListCtrl.SetItem(j, 1, _(u"自动检测"))
            if automated == "True":
                self.test_all_start(0)

        except Exception as e:
            print("line 292: ", e)
            wx.MessageBox(_(u'请选择正确的配置文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return

    def test_start(self, event):
        if isinstance(event, int):
            button_event_id = event - 100
        else:
            button_event_id = event.GetId() - 100  # button ID 100开始

        if self.text_ctrl_json.GetValue():
            pass
        else:
            wx.MessageBox(_(u'请先选择配置文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return

        if self.text_ctrl_py.GetValue():
            if self.port_ctrl_list[button_event_id].GetValue():
                # 点击开始时再初始化
                # self.__port_det(True)  # 检测停止
                self.procese_num += 1
                self.button_start_list[button_event_id].Enable(False)
                self.button_all_start.Enable(False)
                self.message_queue.put(
                    {"id": button_event_id, "msg_id": "PortInit", "PortInfo": self.port_ctrl_list[button_event_id].GetValue()})
            else:
                return None
        else:
            wx.MessageBox(_(u'请先选择测试文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)

    def test_all_start(self, event):
        # self.__port_det(True)  # 检测停止

        if self.text_ctrl_json.GetValue():
            pass
        else:
            wx.MessageBox(_(u'请先选择配置文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return

        if self.text_ctrl_py.GetValue():
            pass
        else:
            wx.MessageBox(_(u'请先选择测试文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return None

        for i in self.port_ctrl_list:
            button_event_id = i.GetId() - 300
            if i.GetValue():
                self.procese_num += 1
                self.button_all_start.Enable(False)
                self.button_start_list[button_event_id].Enable(False)
                self.message_queue.put(
                    {"id": button_event_id, "msg_id": "PortInit",
                     "PortInfo": self.port_ctrl_list[button_event_id].GetValue()})
            else:
                return None

    def port_update(self, arg1):
        print("message:{}".format(arg1))
        self.message_queue.put({"msg_id": "PortUpdate", "PortInfo": arg1})

    def automated_start(self):
        if self.text_ctrl_json.GetValue():
            test_all_start(1)

    def port_update_handler(self, arg1):
        # arg1: {'msg_id': 'PortUpdate', 'PortInfo': ['COM15']}
        # print("message: ", arg1)
        # TODO 设置log按钮状态
        port_list = arg1["PortInfo"][:4] if len(arg1["PortInfo"]) > 4 else arg1["PortInfo"]

        if self._channel_list == []:
            self.button_all_start.Enable(False) if len(port_list) == 0 else self.button_all_start.Enable(True)
        if len(port_list) == 0:
            self.button_all_start.Enable(False)

        if self._channel_list == [] or self._channel_list != port_list:
            self._channel_list = port_list

            exists_port = []
            for i, port_ctrl in enumerate(self.port_ctrl_list):
                if port_ctrl.GetValue() and port_ctrl.GetValue() not in self._channel_list:
                    port_ctrl.SetValue("")
                    self.Imei_ctrl_list[i].SetValue("")
                    self.tips_ctrl_list[i].SetValue("")
                    self.ListCtrl_list[i].DeleteAllItems()
                    self.button_start_list[i].Enable(False)
                elif port_ctrl.GetValue() and port_ctrl.GetValue() in self._channel_list:
                    exists_port.append(port_ctrl.GetValue())

            new_ports = [x for x in self._channel_list if x not in exists_port]
            for new_port in new_ports:
                for port_ctrl in self.port_ctrl_list:
                    if port_ctrl.GetValue() == "":
                        port_ctrl.SetValue(new_port)
                        break

            for i, port_ctrl in enumerate(self.port_ctrl_list):
                if port_ctrl.GetValue() and port_ctrl.GetValue() in new_ports:
                    self.button_start_list[i].Enable(True)
                    if self.text_ctrl_json.GetValue():
                        for j, element in enumerate(self.testFunctions):
                            ListCtrl = self.ListCtrl_list[i]
                            # testFunction = self.testFunctions[j]
                            ListCtrl.InsertItem(j, j)
                            ListCtrl.SetItem(j, 0, self.testMessages[j])
                            if element[-1] == 1:
                                ListCtrl.SetItem(j, 1, _(u"人工检测"))
                            else:
                                ListCtrl.SetItem(j, 1, _(u"自动检测"))
                        print("automated: ", automated)
                        if automated == "True":
                            self.test_start(100+i)


    def port_init_handler(self, arg1):
        ID = arg1["id"]
        # print(ID)
        self.tips_ctrl_list[ID].SetValue("")
        for i in range(len(self.testFunctions)):
            ListCtrl = self.ListCtrl_list[ID]
            ListCtrl.SetItem(i, 2, "")
            ListCtrl.SetItemBackgroundColour(i, (255, 255, 255, 255))

        # 重复测试结果重置
        # ser = serial_handler.SerialHandler(arg1["PortInfo"])
        # imei, iccid = asyncio.run(ser.init_module(arg1["imei"], arg1["iccid"]))
        self.__py_test_fp = open(self.text_ctrl_py.GetValue(), "r", encoding="utf8")
        self.message_queue.put(
            {"id": arg1["id"], "msg_id": "PortTest", "PortInfo": self.port_ctrl_list[arg1["id"]].GetValue(),
             "script": self.__py_test_fp})

    def port_test_handler(self, arg1):
        print("start test", arg1["PortInfo"])
        ID = arg1["id"]
        self.ListCtrl_list[ID].SetItem(0, 2, _(u"测试中"))
        self.ListCtrl_list[ID].SetItemBackgroundColour(0, "Yellow")
        try:
            ser = serial_handler.SerialHandler(arg1["PortInfo"])
            ser.write_module(arg1["script"], self.__exec_py_cmd_list)  # 写入脚本开始测试
            ser.ret_result()
        except Exception as e:
            print("line 406: ", arg1["PortInfo"], e)
            wx.MessageBox(arg1["PortInfo"]+"串口异常，测试脚本写入失败, error %s"%str(e), u'Error', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 1})
            self.button_start_list[arg1["id"]].Enable(True)
            self.procese_num -= 1
            if self.procese_num == 0:
                self.button_all_start.Enable(True)
            return

        try:
            imei = ser.getImei()
            print("imei: ", imei)
            if imei:
                self.Imei_ctrl_list[ID].SetValue(imei)

            ret_result = []
            log = ""
            length = len(self.testFunctions)
            all_result = ""
            for i, testFunction in enumerate(self.testFunctions):
                try:
                    self.ListCtrl_list[ID].SetItem(i, 2, _(u"测试中"))
                    self.ListCtrl_list[ID].SetItemBackgroundColour(i, "Yellow")

                    message = self.testMessages[i]
                    if message == _("版本号检测"):
                        ser.exec_cmd(testFunction[0].replace("\\r\\n", "\r\n"))
                        test_result = ser.ret_result()
                        all_result += test_result
                        print("test_result: ", test_result)
                        version = self.conf.get("version", "version")
                        print("version: ", version)
                        if version in test_result:
                            boolean = ["True"]
                        else:
                            boolean = ["False"]

                    elif message == _("Imei号匹配"):
                        print(os.path.exists(os.path.join(PROJECT_ABSOLUTE_PATH, "imei.txt")))
                        if os.path.exists(os.path.join(PROJECT_ABSOLUTE_PATH, "imei.txt")):
                            f = open(os.path.join(PROJECT_ABSOLUTE_PATH, "imei.txt"))
                            cont = f.read()
                            f.close()
                            print(imei.replace("'","").replace('"',""))
                            if imei.replace("'","").replace('"',"") in cont:
                                boolean = ["True"]
                            else:
                                boolean = ["False"]
                        else:
                            boolean = ["False"]
                    else:
                        cmd = "TestBase." + testFunction[0]
                        ser.exec_cmd(cmd)

                        if self.tips_info.get(message):
                            test_result = ser.ret_result(self.tips_ctrl_list[ID], self.tips_info.get(message))
                        else:
                            test_result = ser.ret_result()  # get recv list
                        all_result += test_result
                        print("test_result: ", test_result)
                        self.logger.write_file(arg1["PortInfo"], test_result)

                        boolean = test_result.split("\r\n")
                        print("boolean: ", boolean)
                        if boolean[-2] == "False":
                            boolean = ["False"]
                        else:
                            if boolean[-2] == "True":
                                boolean = ["True"]
                            elif boolean[-2] == cmd:
                                boolean = ["True"]
                            else:
                                boolean = ["False"]

                except Exception as e:
                    print("line 445: ", e)
                    boolean = ["False"]

                if testFunction[-1] == 1:
                    dlg  = wx.MessageBox(_(u"      当前测试项为:  ") + message + _(u"\r\n\r\n      请确认该测试项 【是否通过】"), self.port_ctrl_list[ID].GetValue(), wx.YES_NO)
                    if dlg == wx.YES:
                        boolean = ["True"]
                    else:
                        boolean = ["False"]

                if boolean == ["False"]:
                    self.ListCtrl_list[ID].SetItem(i, 2, _(u"不通过"))
                    self.ListCtrl_list[ID].SetItemBackgroundColour(i, "Red")
                else:
                    self.ListCtrl_list[ID].SetItem(i, 2, _(u"通过"))
                    self.ListCtrl_list[ID].SetItemBackgroundColour(i, "Green")

                if len(testFunction) == 3 and (isinstance(testFunction[1], int) or isinstance(testFunction[1], float)) and testFunction[1] > 0:
                    time.sleep(testFunction[1])

                ret_result += boolean

            print("all_result: ", all_result)
            self.tips_ctrl_list[ID].SetValue(all_result)
        except Exception as e:
            print("i: ", i)
            print("line 464: ", arg1["PortInfo"], e)
            wx.MessageBox(arg1["PortInfo"]+"测试异常, error %s"%str(e), u'Error', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 1})
            self.button_start_list[arg1["id"]].Enable(True)
            self.procese_num -= 1
            if self.procese_num == 0:
                self.button_all_start.Enable(True)
            return

        if restartSign == "True":
            ser.exec_cmd("from misc import Power\r\nPower.powerRestart()")

        self.button_start_list[arg1["id"]].Enable(True)
        self.procese_num -= 1
        if self.procese_num == 0:
            self.button_all_start.Enable(True)

        # TODO excel写入不要列表形式
        now_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if "False" not in ret_result:
            # self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 1})
            self.__excel_write([arg1["PortInfo"], now_time, imei, "Success", str(list(zip(self.testMessages, ret_result)))])
        else:
            # self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 2})
            self.__excel_write([arg1["PortInfo"], now_time, imei, "Fail", str(list(zip(self.testMessages, ret_result)))])

    # def port_test_end_handler(self, arg1):
    #     self.__port_det(False)

    def __init_excel(self):
        self.__excel_handler = file_handler.ExcelHandler(PROJECT_ABSOLUTE_PATH + "\\Test-Result.xlsx")  # Init Excel
        rows, columns = self.__excel_handler.get_rows_columns()
        if rows == 1 and columns == 1:
            self.__excel_handler.set_cell_value(1, 1, "No.")
            self.__excel_handler.set_cell_value(1, 2, "Com Port")
            self.__excel_handler.set_cell_value(1, 3, "Time")
            self.__excel_handler.set_cell_value(1, 4, "IMEI")
            self.__excel_handler.set_cell_value(1, 5, "Test Result")
            self.__excel_handler.set_cell_value(1, 6, "Test Log")

    def __excel_write(self, result):
        if self.__excel_handler:
            rows, columns = self.__excel_handler.get_rows_columns()
            self.__excel_handler.set_cell_value(rows + 1, 1, rows)
            for i, value in enumerate(result):
                if value == "Success":
                    self.__excel_handler.set_cell_value(rows + 1, i + 2, value, "Green")
                elif value == "Fail":
                    self.__excel_handler.set_cell_value(rows + 1, i + 2, value, "Red")
                else:
                    self.__excel_handler.set_cell_value(rows + 1, i + 2, value)

    def port_process_handler(self):
        while True:
            try:
                message = self.message_queue.get(True, 5)
            except Exception as e:
                message = None
            # try:
            if message:
                msg_id = message.get("msg_id")
                if msg_id == "exit":
                    pass
                elif msg_id == "PortUpdate":
                    # threading.Thread(target=self.port_update_handler, args=(message,)).start()
                    self.port_update_handler(message)
                elif msg_id == "PortInit":
                    # threading.Thread(target=self.port_init_handler, args=(message,)).start()
                    self.port_init_handler(message)
                elif msg_id == "PortTest":
                    threading.Thread(target=self.port_test_handler, args=(message,)).start()
                    # self.port_test_handler(message)
                # elif msg_id == "PortTestEnd":
                    # threading.Thread(target=self.port_test_end_handler, args=(message,)).start()
                    # self.port_test_end_handler(message)
                else:
                    pass
            # except Exception as e:
            #     print("line 546: ", e)

    @staticmethod
    def __port_det(flag):
        if flag:
            tSerialDet.exit_event()
        else:
            tSerialDet.enter_event()

    # end of class FactoryFrame
    def close_window(self, event):
        self.Destroy()
        try:
            self.__excel_handler.close()
        except Exception as e:
            pass
        time.sleep(0.2)
        wx.GetApp().ExitMainLoop()
        wx.Exit()
        process_name = "taskkill /F /IM " + file_name
        p = subprocess.Popen(process_name, shell=True)


class MyApp(wx.App):
    languageTab = locale.getdefaultlocale()[0]
    print("languageTab: ", languageTab)
    # languageTab = "en"
    # 根据系统语言自动设置语言
    # languageTab = "en"
    if languageTab == "zh_CN":
        t = gettext.translation('Chinese', PROJECT_ABSOLUTE_PATH + "\\locale", languages=["zh_CN"])
        t.install()
    elif languageTab == "en":
        t = gettext.translation('English', PROJECT_ABSOLUTE_PATH + "\\locale", languages=["en"])
        t.install()
    else:
        languageTab = "en"
        t = gettext.translation('English', PROJECT_ABSOLUTE_PATH + "\\locale", languages=["en"])
        t.install()

    def __init__(self):
        # print("init-----")
        wx.App.__init__(self, redirect=False, filename="", useBestVisual=True, clearSigInt=True)
        # wx.App.__init__(self, redirect=False, filename=PROJECT_ABSOLUTE_PATH + "\\logs\\software\\run.log",
        #                 useBestVisual=True, clearSigInt=True)
        # print("init finish----")

    def OnInit(self):
        self.factory_frame = FactoryFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.factory_frame)
        self.factory_frame.Show()

        # 串口检测线程
        global tSerialDet
        tSerialDet = serial_list.SerialDetection()
        tSerialDet.setDaemon(True)  # set as deamon, stop thread while main frame exit
        tSerialDet.start()

        set_config(PROJECT_ABSOLUTE_PATH, self.factory_frame)
        return True


class ModuleLog(object):
    """log output file"""
    def __init__(self):
        if not os.path.exists(PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\"):
            os.makedirs(PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\")
        for i in os.walk(PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\"):
            self._log_file_list = i[2]
        if len(self._log_file_list) >= 50:
            os.remove(PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\" + self._log_file_list[0])
        self._log_name = datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".log"

    def write_file(self, port, data):
        try:
            tmp_data = "[" + port + "]\n" + data + "\n"
            with open(PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\" + self._log_name, "a+", encoding="utf-8")as f:
                f.write(tmp_data)
            return True
        except:
            info = sys.exc_info()
            print("write file error.")
            print(info[0], info[1])
            return False


if __name__ == "__main__":
    file_name = os.path.basename(sys.executable)
    app = MyApp()
    app.MainLoop()