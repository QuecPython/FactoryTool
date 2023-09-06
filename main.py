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
                    frame.text_ctrl_py.SetValue(py_file)
                else:
                    conf.set("software", "py_file", "")
                    wx.MessageBox(py_file + _(u"文件不存在"), _(u"提示"), wx.YES_DEFAULT | wx.ICON_INFORMATION)

            if json_file:
                if os.path.exists(py_file):
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


class FactoryFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        kwargs["style"] = kwargs.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwargs)
        self.SetTitle("FactoryTool")
        _icon = wx.NullIcon
        _icon.CopyFromBitmap(wx.Bitmap(PROJECT_ABSOLUTE_PATH + "\\media\\quectel.ico", wx.BITMAP_TYPE_ICO))
        self.SetIcon(_icon)
        self.SetSize((1120, 650))
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
        self.port_imei_list = []
        self.port_iccid_list = []
        for i in range(4):
            ListCtrl = wx.ListCtrl(self.main_panel, 200+i, style=wx.LC_HRULES | wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_VRULES)
            ListCtrl.SetMinSize((-1, 250))
            ListCtrl.AppendColumn(_(u"测项"), format=wx.LIST_FORMAT_LEFT, width=130)
            ListCtrl.AppendColumn(_(u"测试方式"), format=wx.LIST_FORMAT_LEFT, width=70)
            ListCtrl.AppendColumn(_(u"测试结果"), format=wx.LIST_FORMAT_LEFT, width=70)

            button = wx.Button(self.main_panel, 100+i, _("开始"))
            text_ctrl = wx.TextCtrl(self.main_panel, 300+i, "", style=wx.TE_READONLY)
            # text_ctrl_imei = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
            # text_ctrl_iccid = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)

            self.ListCtrl_list.append(ListCtrl)
            self.button_start_list.append(button)
            self.port_ctrl_list.append(text_ctrl)
            # self.port_imei_list.append(text_ctrl_imei)
            # self.port_iccid_list.append(text_ctrl_iccid)

        self.__exec_py_cmd_list = ["from usr.test import TestBase"]
        self.__init_excel()

        # event message queue
        self.message_queue = Queue(maxsize=10)
        self._channel_list = []
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

        self.__set_properties()
        self.__do_layout()

    def __set_properties(self):
        for i in range(4):
            self.button_start_list[i].SetMinSize((150, 50))
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
            sizer_6.Add(self.port_ctrl_list[i], 0, 0, 1)
            sizer_6.Add(self.ListCtrl_list[i], 1, 1, 1)
            sizer_6.Add((0, 10), 0, 0, 0)
            sizer_6.Add(self.button_start_list[i], 0, 1, 0)
            sizer_5.Add(sizer_6, 1, 0, 0)
        sizer_2.Add(sizer_5, 0, wx.EXPAND, 0)

        sizer_2.Add((0, 30), 0, 1, 1)
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
        event.Skip()

    def initConfigFile(self):
        self.conf.add_section("software")
        self.conf.set("software", "py_file", "")
        self.conf.set("software", "json_file", "")
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
                self.testFunctions = [i[1:] for i in info]
                self.testMessages = [i[0] for i in info]
            self.text_ctrl_json.SetValue(json_file)
            self.conf.set("software", "json_file", json_file)
            with open(PROJECT_ABSOLUTE_PATH + "\\config.ini", "w+", encoding='utf-8') as f:
                self.conf.write(f)

            for i, port_ctrl in enumerate(self.port_ctrl_list):
                if port_ctrl.GetValue():
                    self.ListCtrl_list[i].DeleteAllItems()
                    for j, element in enumerate(self.testFunctions):
                        ListCtrl = self.ListCtrl_list[i]
                        # testFunction = self.testFunctions[j]
                        ListCtrl.InsertItem(j, j)
                        ListCtrl.SetItem(j, 0, self.testMessages[j])
                        if element[1] == 1:
                            ListCtrl.SetItem(j, 1, _(u"人工检测"))
                        else:
                            ListCtrl.SetItem(j, 1, _(u"自动检测"))
        except Exception as e:
            print(e)
            wx.MessageBox(_(u'请选择正确的配置文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return

    def test_start(self, event):
        button_event_id = event.GetId() - 100  # button ID 100开始

        if self.text_ctrl_json.GetValue():
            pass
        else:
            wx.MessageBox(_(u'请先选择配置文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return

        if self.text_ctrl_py.GetValue():
            if self.port_ctrl_list[button_event_id].GetValue():
                # 点击开始时再初始化
                self.__port_det(True)  # 检测停止

                self.procese_num += 1
                self.button_all_start.Enable(False)
                self.button_start_list[button_event_id].Enable(False)
                self.message_queue.put(
                    {"id": button_event_id, "msg_id": "PortInit", "PortInfo": self.port_ctrl_list[button_event_id].GetValue()})
            else:
                return None
        else:
            wx.MessageBox(_(u'请先选择测试文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)

    def test_all_start(self, event):
        self.__port_det(True)  # 检测停止

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

    def port_update_handler(self, arg1):
        # TODO 设置log按钮状态
        port_list = arg1["PortInfo"][:4] if len(arg1["PortInfo"]) > 4 else arg1["PortInfo"]
        if len(port_list) < 4:
            self.button_all_start.Enable(False) if len(port_list) == 0 else self.button_all_start.Enable(True)
            port_list = port_list + [""] * (4 - len(port_list))
        if self._channel_list == [] or self._channel_list != port_list:
            self._channel_list = port_list
            for port_ctrl in self.port_ctrl_list:
                port_ctrl.SetValue("")
            for ListCtrl in self.ListCtrl_list:
                ListCtrl.DeleteAllItems()

            for i, element in enumerate(self.button_start_list):
                if self._channel_list[i]:
                    self.port_ctrl_list[i].SetValue(self._channel_list[i])
                    self.button_start_list[i].Enable(True)

                    if self.text_ctrl_json.GetValue():
                        for j, element in enumerate(self.testFunctions):
                            ListCtrl = self.ListCtrl_list[i]
                            # testFunction = self.testFunctions[j]
                            ListCtrl.InsertItem(j, j)
                            ListCtrl.SetItem(j, 0, self.testMessages[j])
                            if element[1] == 1:
                                ListCtrl.SetItem(j, 1, _(u"人工检测"))
                            else:
                                ListCtrl.SetItem(j, 1, _(u"自动检测"))
                else:
                    self.button_start_list[i].Enable(False)

    def port_init_handler(self, arg1):
        ID = arg1["id"]
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
        print("start test")
        ID = arg1["id"]
        try:
            ser = serial_handler.SerialHandler(arg1["PortInfo"])
            ser.write_module(arg1["script"], self.__exec_py_cmd_list)  # 写入脚本开始测试
            ser.ret_result()
        except Exception as e:
            print(arg1["PortInfo"], e)
            wx.MessageBox(arg1["PortInfo"]+"串口异常，测试脚本写入失败, error %s"%str(e), u'Error', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 1})
            self.button_start_list[arg1["id"]].Enable(True)
            self.procese_num -= 1
            if self.procese_num == 0:
                self.button_all_start.Enable(True)
            return

        ret_result = []
        log = ""
        length = len(self.testFunctions)
        for i, testFunction in enumerate(self.testFunctions):
            try:
                self.ListCtrl_list[ID].SetItem(i, 2, _(u"测试中"))
                self.ListCtrl_list[ID].SetItemBackgroundColour(i, "Yellow")

                message = self.testMessages[i]

                cmd = "TestBase." + testFunction[0] + "()"
                ser.exec_cmd(cmd)

                test_result = ser.ret_result()  # get recv list
                self.logger.write_file(arg1["PortInfo"], test_result)

                boolean = test_result.split("\r\n")

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
                print(e)
                boolean = ["False"]

            if testFunction[1] == 1:
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

            ret_result += boolean

        test_results = []
        for i in ret_result:
            if i == "True":
                test_results.append(_(u"通过"))
            elif i == "False":
                test_results.append(_(u"不通过"))
            else:
                test_results.append(i)

        test_method = []
        for i in self.testFunctions:
            if i[1] == 1:
                test_method.append(_(u"人工检测"))
            else:
                test_method.append(_(u"自动检测"))

        # TODO excel写入不要列表形式
        if "False" not in ret_result:
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 1})
            self.__excel_write([arg1["PortInfo"], "Success", str(list(zip(self.testMessages, ret_result)))])
        else:
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 2})
            self.__excel_write([arg1["PortInfo"], "Fail", str(list(zip(self.testMessages, ret_result)))])
        self.button_start_list[arg1["id"]].Enable(True)
        self.procese_num -= 1
        if self.procese_num == 0:
            self.button_all_start.Enable(True)

    def port_test_end_handler(self, arg1):
        self.__port_det(False)

    def __init_excel(self):
        self.__excel_handler = file_handler.ExcelHandler(PROJECT_ABSOLUTE_PATH + "\\Test-Result.xlsx")  # Init Excel
        rows, columns = self.__excel_handler.get_rows_columns()
        if rows == 1 and columns == 1:
            self.__excel_handler.set_cell_value(1, 1, "No.")
            self.__excel_handler.set_cell_value(1, 2, "Com Port")
            self.__excel_handler.set_cell_value(1, 3, "Test Result")
            self.__excel_handler.set_cell_value(1, 4, "Test Log")

    def __excel_write(self, result):
        if self.__excel_handler:
            rows, columns = self.__excel_handler.get_rows_columns()
            self.__excel_handler.set_cell_value(rows + 1, 1, rows)
            for i, value in enumerate(result):
                self.__excel_handler.set_cell_value(rows + 1, i + 2, value)

    def port_process_handler(self):
        while True:
            try:
                message = self.message_queue.get(True, 5)
            except Exception as e:
                message = None
            try:
                if message:
                    # print("message:{}".format(message))
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
                    elif msg_id == "PortTestEnd":
                        threading.Thread(target=self.port_test_end_handler, args=(message,)).start()
                        # self.port_test_end_handler(message)
                    else:
                        pass
            except Exception as e:
                print(e)

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