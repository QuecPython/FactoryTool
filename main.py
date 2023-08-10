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
from log_redirect import RedirectErr, RedirectStd
from pubsub import pub
from queue import Queue
import codecs
import json

# PROJECT_ABSOLUTE_PATH = os.path.dirname(os.path.abspath(__file__))

if hasattr(sys, '_MEIPASS'):
    PROJECT_ABSOLUTE_PATH = os.path.dirname(os.path.realpath(sys.executable))
else:
    PROJECT_ABSOLUTE_PATH = os.path.dirname(os.path.abspath(__file__))

class FactoryFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        kwargs["style"] = kwargs.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwargs)
        self.SetSize((876, 600))
        self.Center()
        # stdout log
        # sys.stderr = RedirectErr(self, PROJECT_ABSOLUTE_PATH)
        # sys.stdout = RedirectStd(self, PROJECT_ABSOLUTE_PATH)
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
        wxg_tmp_menu_1.Append(2002, _(u"工具日志"), "")
        self.Bind(wx.EVT_MENU, self.__menu_handler, id=2002)
        self.menuBar.Append(wxg_tmp_menu_1, _(u"日志"))
        self.SetMenuBar(self.menuBar)
        # Menu Bar end
        self.statusBar = self.CreateStatusBar(3)
        self.main_panel = wx.Panel(self, wx.ID_ANY)
        self.load_py = wx.Button(self.main_panel, wx.ID_ANY, _("选择"))
        self.text_ctrl_py = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.load_json = wx.Button(self.main_panel, wx.ID_ANY, _("选择"))
        self.text_ctrl_json = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_3 = wx.Button(self.main_panel, wx.ID_ANY, _("全部开始"), style=wx.BU_TOP)
        self.button_4 = wx.Button(self.main_panel, 100, _("开始"))
        self.text_ctrl_7 = wx.TextCtrl(self.main_panel, 10, "", style=wx.TE_READONLY)
        self.text_ctrl_8 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_9 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_10 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_11 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_2 = wx.Button(self.main_panel, 200, _("日志"))

        self.button_5 = wx.Button(self.main_panel, 101, _("开始"))
        self.text_ctrl_13 = wx.TextCtrl(self.main_panel, 11, "", style=wx.TE_READONLY)
        self.text_ctrl_14 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_15 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_16 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_17 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_12 = wx.Button(self.main_panel, 201, _("日志"))

        self.button_6 = wx.Button(self.main_panel, 102, _("开始"))
        self.text_ctrl_19 = wx.TextCtrl(self.main_panel, 12, "", style=wx.TE_READONLY)
        self.text_ctrl_20 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_21 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_22 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_23 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_13 = wx.Button(self.main_panel, 202, _("日志"))

        self.button_7 = wx.Button(self.main_panel, 103, _("开始"))
        self.text_ctrl_25 = wx.TextCtrl(self.main_panel, 13, "", style=wx.TE_READONLY)
        self.text_ctrl_26 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_27 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_28 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_29 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_14 = wx.Button(self.main_panel, 203, _("日志"))

        self.button_8 = wx.Button(self.main_panel, 104, _("开始"))
        self.text_ctrl_31 = wx.TextCtrl(self.main_panel, 14, "", style=wx.TE_READONLY)
        self.text_ctrl_32 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_33 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_34 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_35 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_15 = wx.Button(self.main_panel, 204, _("日志"))

        self.button_9 = wx.Button(self.main_panel, 105, _("开始"))
        self.text_ctrl_37 = wx.TextCtrl(self.main_panel, 15, "", style=wx.TE_READONLY)
        self.text_ctrl_38 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_39 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_40 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_41 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_16 = wx.Button(self.main_panel, 205, _("日志"))

        self.button_10 = wx.Button(self.main_panel, 106, _("开始"))
        self.text_ctrl_43 = wx.TextCtrl(self.main_panel, 16, "", style=wx.TE_READONLY)
        self.text_ctrl_44 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_45 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_46 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_47 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_17 = wx.Button(self.main_panel, 206, _("日志"))

        self.button_11 = wx.Button(self.main_panel, 107, _("开始"))
        self.text_ctrl_49 = wx.TextCtrl(self.main_panel, 17, "", style=wx.TE_READONLY)
        self.text_ctrl_50 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_51 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_52 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.text_ctrl_53 = wx.TextCtrl(self.main_panel, wx.ID_ANY, "", style=wx.TE_READONLY)
        self.button_18 = wx.Button(self.main_panel, 207, _("日志"))

        # button list
        self.button_start_list = (self.button_4, self.button_5, self.button_6, self.button_7, self.button_8, self.button_9, self.button_10, self.button_11)
        self.button_log_list = (self.button_2, self.button_12, self.button_13, self.button_14, self.button_15, self.button_16, self.button_17, self.button_18)
        [i.Enable(False) for i in self.button_start_list]
        [i.Enable(False) for i in self.button_log_list]
        # text_ctrl list
        self.port_ctrl_list = (self.text_ctrl_7, self.text_ctrl_13, self.text_ctrl_19, self.text_ctrl_25,
                               self.text_ctrl_31, self.text_ctrl_37, self.text_ctrl_43, self.text_ctrl_49)
        self.port_imei_list = (self.text_ctrl_8, self.text_ctrl_14, self.text_ctrl_20, self.text_ctrl_26,
                               self.text_ctrl_32, self.text_ctrl_38, self.text_ctrl_44, self.text_ctrl_50)
        self.port_iccid_list = (self.text_ctrl_9, self.text_ctrl_15, self.text_ctrl_21, self.text_ctrl_27,
                                self.text_ctrl_33, self.text_ctrl_39, self.text_ctrl_45, self.text_ctrl_51)
        self.port_time_list = (self.text_ctrl_10, self.text_ctrl_16, self.text_ctrl_22, self.text_ctrl_28,
                               self.text_ctrl_34, self.text_ctrl_40, self.text_ctrl_46, self.text_ctrl_52)
        self.port_result_list = (self.text_ctrl_11, self.text_ctrl_17, self.text_ctrl_23, self.text_ctrl_29,
                                 self.text_ctrl_35, self.text_ctrl_41, self.text_ctrl_47, self.text_ctrl_53)
        self.__py_exec_result = list(" ") * 8
        self.__py_test_fp = None  # 测试的脚本文件
        self.__exec_py_cmd_list = ["from usr.test import TestBase"]
        self.__test_result = [0 for i in range(8)]  # 测试结果
        self.statusBarTimer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.__status_bar_timer_fresh, self.statusBarTimer)
        self.statusBarTimer.Start(100)
        pub.subscribe(self.__update_status_bar, "statusBarUpdate")

        # event message queue
        self.message_queue = Queue(maxsize=10)
        self._channel_list = []
        threading.Thread(target=self.port_process_handler, args=()).start()

        # bind Func
        self.Bind(wx.EVT_BUTTON, self.__load_py_file, self.load_py)
        self.Bind(wx.EVT_BUTTON, self.__load_json_file, self.load_json)
        self.Bind(wx.EVT_BUTTON, self.test_all_start, self.button_3)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_4)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_5)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_6)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_7)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_8)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_9)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_10)
        self.Bind(wx.EVT_BUTTON, self.test_start, self.button_11)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_2)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_12)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_13)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_14)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_15)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_16)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_17)
        self.Bind(wx.EVT_BUTTON, self.log_terminal, self.button_18)
        self.Bind(wx.EVT_CLOSE, self.close_window)

        # bind Timer
        self.statusBarTimer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.statusbar_fresh, self.statusBarTimer)
        self.statusBarTimer.Start(100)

        # pub message
        pub.subscribe(self.port_update, "serialUpdate")

        self.__set_properties()
        self.__do_layout()

    def __set_properties(self):
        self.SetTitle("Factory Tool")
        # _icon = wx.NullIcon
        # _icon.CopyFromBitmap(wx.Bitmap(PROJECT_ABSOLUTE_PATH + "\\images\\quectel.ico", wx.BITMAP_TYPE_ICO))
        # self.SetIcon(_icon)
        self.SetBackgroundColour(wx.NullColour)
        self.SetForegroundColour(wx.Colour(0, 0, 0))
        self.statusBar.SetStatusWidths([200, -1, 155])

        # statusbar fields
        statusBar_fields = [_(u"  Welcome to QuecPython"), _(u" Current status: ready..."), " 0000-00-00 00:00:00"]
        for i in range(len(statusBar_fields)):
            self.statusBar.SetStatusText(statusBar_fields[i], i)

        self.load_py.SetMinSize((80, 25))
        self.load_json.SetMinSize((80, 25))
        self.text_ctrl_7.SetMinSize((74, 25))
        self.text_ctrl_8.SetMinSize((117, 25))
        self.text_ctrl_9.SetMinSize((123, 25))
        self.text_ctrl_11.SetMinSize((148, 25))
        self.button_2.SetMinSize((98, 25))
        self.text_ctrl_13.SetMinSize((74, 25))
        self.text_ctrl_14.SetMinSize((117, 25))
        self.text_ctrl_15.SetMinSize((123, 25))
        self.text_ctrl_17.SetMinSize((148, 25))
        self.button_12.SetMinSize((98, 25))
        self.text_ctrl_19.SetMinSize((74, 25))
        self.text_ctrl_20.SetMinSize((117, 25))
        self.text_ctrl_21.SetMinSize((123, 25))
        self.text_ctrl_23.SetMinSize((148, 25))
        self.button_13.SetMinSize((98, 25))
        self.text_ctrl_25.SetMinSize((74, 25))
        self.text_ctrl_26.SetMinSize((117, 25))
        self.text_ctrl_27.SetMinSize((123, 25))
        self.text_ctrl_29.SetMinSize((148, 25))
        self.button_14.SetMinSize((98, 25))
        self.text_ctrl_31.SetMinSize((74, 25))
        self.text_ctrl_32.SetMinSize((117, 25))
        self.text_ctrl_33.SetMinSize((123, 25))
        self.text_ctrl_35.SetMinSize((148, 25))
        self.button_15.SetMinSize((98, 25))
        self.text_ctrl_37.SetMinSize((74, 25))
        self.text_ctrl_38.SetMinSize((117, 25))
        self.text_ctrl_39.SetMinSize((123, 25))
        self.text_ctrl_41.SetMinSize((148, 25))
        self.button_16.SetMinSize((98, 25))
        self.text_ctrl_43.SetMinSize((74, 25))
        self.text_ctrl_44.SetMinSize((117, 25))
        self.text_ctrl_45.SetMinSize((123, 25))
        self.text_ctrl_47.SetMinSize((148, 25))
        self.button_17.SetMinSize((98, 25))
        self.text_ctrl_49.SetMinSize((74, 25))
        self.text_ctrl_50.SetMinSize((117, 25))
        self.text_ctrl_51.SetMinSize((123, 25))
        self.text_ctrl_53.SetMinSize((148, 25))
        self.button_18.SetMinSize((98, 25))

    def __do_layout(self):
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_5 = wx.StaticBoxSizer(wx.StaticBox(self.main_panel, wx.ID_ANY, _("工厂测试")), wx.VERTICAL)
        sizer_14 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_13 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_12 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_11 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_10 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_9 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_4 = wx.StaticBoxSizer(wx.StaticBox(self.main_panel, wx.ID_ANY, _("选择配置文件")), wx.HORIZONTAL)
        sizer_3 = wx.StaticBoxSizer(wx.StaticBox(self.main_panel, wx.ID_ANY, _("选择测试脚本")), wx.HORIZONTAL)
        sizer_1.Add((20, 4), 0, 0, 0)
        sizer_2.Add((20, 4), 0, 0, 0)
        sizer_3.Add((4, 20), 0, 0, 0)
        sizer_3.Add(self.load_py, 0, 0, 0)
        sizer_3.Add((20, 20), 0, 0, 0)
        sizer_3.Add(self.text_ctrl_py, 1, 0, 0)
        sizer_3.Add((20, 20), 0, 0, 0)
        sizer_2.Add(sizer_3, 0, wx.EXPAND, 0)
        sizer_4.Add((4, 20), 0, 0, 0)
        sizer_4.Add(self.load_json, 0, 0, 0)
        sizer_4.Add((20, 20), 0, 0, 0)
        sizer_4.Add(self.text_ctrl_json, 1, 0, 0)
        sizer_4.Add((20, 20), 0, 0, 0)
        sizer_2.Add((20, 4), 0, 0, 0)
        sizer_2.Add(sizer_4, 0, wx.EXPAND, 0)
        sizer_6.Add((20, 20), 0, 0, 0)
        sizer_6.Add(self.button_3, 0, 0, 0)
        sizer_6.Add((20, 20), 0, 0, 0)
        label_9 = wx.StaticText(self.main_panel, wx.ID_ANY, _(u"端口"))
        sizer_6.Add(label_9, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_6.Add((60, 20), 0, 0, 0)
        label_10 = wx.StaticText(self.main_panel, wx.ID_ANY, _("IMEI"))
        sizer_6.Add(label_10, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_6.Add((100, 20), 0, 0, 0)
        label_11 = wx.StaticText(self.main_panel, wx.ID_ANY, _("ICCID"))
        sizer_6.Add(label_11, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_6.Add((100, 20), 0, 0, 0)
        label_12 = wx.StaticText(self.main_panel, wx.ID_ANY, _(u"测试时间"))
        sizer_6.Add(label_12, 1, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_6.Add((10, 20), 0, 0, 0)
        label_13 = wx.StaticText(self.main_panel, wx.ID_ANY, _(u"测试结果"))
        sizer_6.Add(label_13, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_6.Add((100, 20), 0, 0, 0)
        label_14 = wx.StaticText(self.main_panel, wx.ID_ANY, _(u"查看日志"))
        sizer_6.Add(label_14, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_6.Add((50, 20), 0, 0, 0)
        sizer_5.Add(sizer_6, 0, wx.EXPAND | wx.TOP, 0)
        sizer_5.Add((20, 5), 0, 0, 0)
        label_15 = wx.StaticText(self.main_panel, wx.ID_ANY, "1", style=wx.ALIGN_CENTER)
        label_15.SetMinSize((20, 18))
        sizer_7.Add(label_15, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_7.Add(self.button_4, 0, 0, 0)
        sizer_7.Add((20, 20), 0, 0, 0)
        sizer_7.Add(self.text_ctrl_7, 0, wx.ALL, 0)
        sizer_7.Add((10, 20), 0, 0, 0)
        sizer_7.Add(self.text_ctrl_8, 0, 0, 0)
        sizer_7.Add((10, 20), 0, 0, 0)
        sizer_7.Add(self.text_ctrl_9, 0, 0, 0)
        sizer_7.Add((10, 20), 0, 0, 0)
        sizer_7.Add(self.text_ctrl_10, 1, 0, 0)
        sizer_7.Add((10, 20), 0, 0, 0)
        sizer_7.Add(self.text_ctrl_11, 0, 0, 0)
        sizer_7.Add((10, 20), 0, 0, 0)
        sizer_7.Add(self.button_2, 0, 0, 0)
        sizer_7.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_7, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_16 = wx.StaticText(self.main_panel, wx.ID_ANY, "2", style=wx.ALIGN_CENTER)
        label_16.SetMinSize((20, 18))
        sizer_8.Add(label_16, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_8.Add(self.button_5, 0, 0, 0)
        sizer_8.Add((20, 20), 0, 0, 0)
        sizer_8.Add(self.text_ctrl_13, 0, wx.ALL, 0)
        sizer_8.Add((10, 20), 0, 0, 0)
        sizer_8.Add(self.text_ctrl_14, 0, 0, 0)
        sizer_8.Add((10, 20), 0, 0, 0)
        sizer_8.Add(self.text_ctrl_15, 0, 0, 0)
        sizer_8.Add((10, 20), 0, 0, 0)
        sizer_8.Add(self.text_ctrl_16, 1, 0, 0)
        sizer_8.Add((10, 20), 0, 0, 0)
        sizer_8.Add(self.text_ctrl_17, 0, 0, 0)
        sizer_8.Add((10, 20), 0, 0, 0)
        sizer_8.Add(self.button_12, 0, 0, 0)
        sizer_8.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_8, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_17 = wx.StaticText(self.main_panel, wx.ID_ANY, "3", style=wx.ALIGN_CENTER)
        label_17.SetMinSize((20, 18))
        sizer_9.Add(label_17, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_9.Add(self.button_6, 0, 0, 0)
        sizer_9.Add((20, 20), 0, 0, 0)
        sizer_9.Add(self.text_ctrl_19, 0, wx.ALL, 0)
        sizer_9.Add((10, 20), 0, 0, 0)
        sizer_9.Add(self.text_ctrl_20, 0, 0, 0)
        sizer_9.Add((10, 20), 0, 0, 0)
        sizer_9.Add(self.text_ctrl_21, 0, 0, 0)
        sizer_9.Add((10, 20), 0, 0, 0)
        sizer_9.Add(self.text_ctrl_22, 1, 0, 0)
        sizer_9.Add((10, 20), 0, 0, 0)
        sizer_9.Add(self.text_ctrl_23, 0, 0, 0)
        sizer_9.Add((10, 20), 0, 0, 0)
        sizer_9.Add(self.button_13, 0, 0, 0)
        sizer_9.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_9, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_18 = wx.StaticText(self.main_panel, wx.ID_ANY, "4", style=wx.ALIGN_CENTER)
        label_18.SetMinSize((20, 18))
        sizer_10.Add(label_18, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_10.Add(self.button_7, 0, 0, 0)
        sizer_10.Add((20, 20), 0, 0, 0)
        sizer_10.Add(self.text_ctrl_25, 0, wx.ALL, 0)
        sizer_10.Add((10, 20), 0, 0, 0)
        sizer_10.Add(self.text_ctrl_26, 0, 0, 0)
        sizer_10.Add((10, 20), 0, 0, 0)
        sizer_10.Add(self.text_ctrl_27, 0, 0, 0)
        sizer_10.Add((10, 20), 0, 0, 0)
        sizer_10.Add(self.text_ctrl_28, 1, 0, 0)
        sizer_10.Add((10, 20), 0, 0, 0)
        sizer_10.Add(self.text_ctrl_29, 0, 0, 0)
        sizer_10.Add((10, 20), 0, 0, 0)
        sizer_10.Add(self.button_14, 0, 0, 0)
        sizer_10.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_10, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_19 = wx.StaticText(self.main_panel, wx.ID_ANY, "5", style=wx.ALIGN_CENTER)
        label_19.SetMinSize((20, 18))
        sizer_11.Add(label_19, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_11.Add(self.button_8, 0, 0, 0)
        sizer_11.Add((20, 20), 0, 0, 0)
        sizer_11.Add(self.text_ctrl_31, 0, wx.ALL, 0)
        sizer_11.Add((10, 20), 0, 0, 0)
        sizer_11.Add(self.text_ctrl_32, 0, 0, 0)
        sizer_11.Add((10, 20), 0, 0, 0)
        sizer_11.Add(self.text_ctrl_33, 0, 0, 0)
        sizer_11.Add((10, 20), 0, 0, 0)
        sizer_11.Add(self.text_ctrl_34, 1, 0, 0)
        sizer_11.Add((10, 20), 0, 0, 0)
        sizer_11.Add(self.text_ctrl_35, 0, 0, 0)
        sizer_11.Add((10, 20), 0, 0, 0)
        sizer_11.Add(self.button_15, 0, 0, 0)
        sizer_11.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_11, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_20 = wx.StaticText(self.main_panel, wx.ID_ANY, "6", style=wx.ALIGN_CENTER)
        label_20.SetMinSize((20, 18))
        sizer_12.Add(label_20, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_12.Add(self.button_9, 0, 0, 0)
        sizer_12.Add((20, 20), 0, 0, 0)
        sizer_12.Add(self.text_ctrl_37, 0, wx.ALL, 0)
        sizer_12.Add((10, 20), 0, 0, 0)
        sizer_12.Add(self.text_ctrl_38, 0, 0, 0)
        sizer_12.Add((10, 20), 0, 0, 0)
        sizer_12.Add(self.text_ctrl_39, 0, 0, 0)
        sizer_12.Add((10, 20), 0, 0, 0)
        sizer_12.Add(self.text_ctrl_40, 1, 0, 0)
        sizer_12.Add((10, 20), 0, 0, 0)
        sizer_12.Add(self.text_ctrl_41, 0, 0, 0)
        sizer_12.Add((10, 20), 0, 0, 0)
        sizer_12.Add(self.button_16, 0, 0, 0)
        sizer_12.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_12, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_21 = wx.StaticText(self.main_panel, wx.ID_ANY, "7", style=wx.ALIGN_CENTER)
        label_21.SetMinSize((20, 18))
        sizer_13.Add(label_21, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_13.Add(self.button_10, 0, 0, 0)
        sizer_13.Add((20, 20), 0, 0, 0)
        sizer_13.Add(self.text_ctrl_43, 0, wx.ALL, 0)
        sizer_13.Add((10, 20), 0, 0, 0)
        sizer_13.Add(self.text_ctrl_44, 0, 0, 0)
        sizer_13.Add((10, 20), 0, 0, 0)
        sizer_13.Add(self.text_ctrl_45, 0, 0, 0)
        sizer_13.Add((10, 20), 0, 0, 0)
        sizer_13.Add(self.text_ctrl_46, 1, 0, 0)
        sizer_13.Add((10, 20), 0, 0, 0)
        sizer_13.Add(self.text_ctrl_47, 0, 0, 0)
        sizer_13.Add((10, 20), 0, 0, 0)
        sizer_13.Add(self.button_17, 0, 0, 0)
        sizer_13.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_13, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        label_22 = wx.StaticText(self.main_panel, wx.ID_ANY, "8", style=wx.ALIGN_CENTER)
        label_22.SetMinSize((20, 18))
        sizer_14.Add(label_22, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 0)
        sizer_14.Add(self.button_11, 0, 0, 0)
        sizer_14.Add((20, 20), 0, 0, 0)
        sizer_14.Add(self.text_ctrl_49, 0, wx.ALL, 0)
        sizer_14.Add((10, 20), 0, 0, 0)
        sizer_14.Add(self.text_ctrl_50, 0, 0, 0)
        sizer_14.Add((10, 20), 0, 0, 0)
        sizer_14.Add(self.text_ctrl_51, 0, 0, 0)
        sizer_14.Add((10, 20), 0, 0, 0)
        sizer_14.Add(self.text_ctrl_52, 1, 0, 0)
        sizer_14.Add((10, 20), 0, 0, 0)
        sizer_14.Add(self.text_ctrl_53, 0, 0, 0)
        sizer_14.Add((10, 20), 0, 0, 0)
        sizer_14.Add(self.button_18, 0, 0, 0)
        sizer_14.Add((10, 20), 0, 0, 0)
        sizer_5.Add(sizer_14, 0, wx.ALL | wx.EXPAND, 0)
        sizer_5.Add((20, 20), 0, 0, 0)
        sizer_2.Add(sizer_5, 0, wx.ALL | wx.EXPAND, 1)
        self.main_panel.SetSizer(sizer_2)
        sizer_1.Add(self.main_panel, 1, wx.EXPAND | wx.TOP, 0)
        self.SetSizer(sizer_1)
        self.Layout()

    def __status_bar_timer_fresh(self, event):
        now = datetime.datetime.now()
        self.statusBar.SetStatusText(" " + now.strftime("%Y-%m-%d %H:%M:%S"), 2)

    def __update_status_bar(self, arg1):
        self.statusBar.SetStatusText(arg1[0], arg1[1])

    def __menu_handler(self, event):  # wxGlade: FactoryFrame.<event_handler>
        if event.GetId() == 1001:
            DialogControl(self, "demo").Show()

        elif event.GetId() == 1002:
            file_path, app_path = "notepad.exe", PROJECT_ABSOLUTE_PATH + "\\module_test.py"
            raw_file_path = r'%s' % file_path
            raw_app_path = r'%s' % app_path
            subprocess.call([raw_file_path, raw_app_path])
        elif event.GetId() == 1003:
            excel_file = subprocess.Popen(["start", "/WAIT", PROJECT_ABSOLUTE_PATH + "Test-Result.xlsx"], shell=True)
            # psutil.Process(excel_file.pid).get_children()[0].kill()
            excel_file.poll()
        elif event.GetId() == 2001:
            p = subprocess.Popen("explorer.exe " + PROJECT_ABSOLUTE_PATH + "\\logs\\software\\", shell=True)
        elif event.GetId() == 2002:
            p = subprocess.Popen("explorer.exe " + PROJECT_ABSOLUTE_PATH + "\\logs\\apps\\", shell=True)
        event.Skip()

    def __load_py_file(self, event):
        defDir, defFile = '', ''  # default dir/ default file
        dlg = wx.FileDialog(self, u'Open Py File', defDir, defFile, 'Python file (*.py)|*.py',
                            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return
        py_cmd = file_handler.ScriptHandler(dlg.GetPath())
        if py_cmd.script_check():
            self.text_ctrl_py.SetValue(dlg.GetPath())
        else:
            wx.MessageBox(_(u'py代码有语法错误'), u'Warn', wx.YES_DEFAULT | wx.ICON_INFORMATION)

    def __load_json_file(self, event):
        print("json button is press")
        defDir, defFile = '', ''  # default dir/ default file
        dlg = wx.FileDialog(self, u'Open Config File', defDir, defFile, 'Config file (*)|*',
                            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return
        self.text_ctrl_json.SetValue(dlg.GetPath())
        # TODO ymodem处理文件

    def test_start(self, event):
        button_event_id = event.GetId() - 100  # button ID 100开始
        self.button_start_list[button_event_id].SetLabel(_("测试中"))
        self.button_start_list[button_event_id].Enable(False)
        jsonName = PROJECT_ABSOLUTE_PATH + "\\sort_setting.json"
        with codecs.open(jsonName, 'r', 'utf-8') as f:
            data = json.load(f)
            self.testFunctions = data["sort"]
            self.testMessages = data["message"]

        if self.text_ctrl_py.GetValue():
            if self.port_ctrl_list[button_event_id].GetValue():
                # 点击开始时再初始化
                self.__port_det(True)  # 检测停止
                self.message_queue.put(
                    {"id": button_event_id, "msg_id": "PortInit", "PortInfo": self.port_ctrl_list[button_event_id].GetValue(),
                     "imei": self.port_imei_list[button_event_id], "iccid": self.port_iccid_list[button_event_id]})
            else:
                return None
        else:
            wx.MessageBox(_(u'请先选择测试文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)

    def test_all_start(self, event):
        self.__port_det(True)  # 检测停止
        if self.text_ctrl_py.GetValue():
            pass
        else:
            wx.MessageBox(_(u'请先选择测试文件'), u'提示', wx.YES_DEFAULT | wx.ICON_INFORMATION)
            return None
        for i in self.port_ctrl_list:
            button_event_id = i.GetId() - 10
            if i.GetValue():
                self.message_queue.put(
                    {"id": button_event_id, "msg_id": "PortInit",
                     "PortInfo": self.port_ctrl_list[button_event_id].GetValue(),
                     "imei": self.port_imei_list[button_event_id], "iccid": self.port_iccid_list[button_event_id]})
                # time.sleep(2)
                # self.message_queue.put(
                #     {"id": button_event_id, "msg_id": "PortTest",
                #      "PortInfo": self.port_ctrl_list[button_event_id].GetValue(),
                #      "script": self.__py_test_fp, "time": self.port_time_list[button_event_id],
                #      "result": self.port_result_list[button_event_id]})
            else:
                return None

    def log_terminal(self, event):
        event_id = event.GetId()
        # print(self.port_ctrl_list[event_id-200].GetValue(), self.__py_exec_result[event_id-200])
        dia = DialogControl(self, self.port_ctrl_list[event_id-200].GetValue(), self.__py_exec_result[event_id-200])
        #
        res = dia.Show()

    def statusbar_fresh(self, event):
        # self.statusBar.SetStatusText(u" 当前状态: 就绪...", 1)
        now = datetime.datetime.now()
        self.statusBar.SetStatusText(" " + now.strftime("%Y-%m-%d %H:%M:%S"), 2)

    def port_update(self, arg1):
        # print("message:{}".format(arg1))
        self.message_queue.put({"msg_id": "PortUpdate", "PortInfo": arg1})

    def port_update_handler(self, arg1):
        # TODO 设置log按钮状态
        port_list = arg1["PortInfo"][:8] if len(arg1["PortInfo"]) > 8 else arg1["PortInfo"]
        if len(port_list) < 8:
            self.button_3.Enable(False) if len(port_list) == 0 else self.button_3.Enable(True)
            port_list = port_list + [""] * (8 - len(port_list))
        if self._channel_list == [] or self._channel_list != port_list:
            self._channel_list = port_list
            for j in [self.port_ctrl_list, self.port_imei_list, self.port_iccid_list, self.port_time_list,
                      self.port_result_list]:
                [i.SetValue("") for i in j]
            for i, element in enumerate(self.button_start_list):
                self.button_log_list[i].Enable(False)
                # self.button_next_list[i].Enable(False)
                if self._channel_list[i]:
                    self.port_ctrl_list[i].SetValue(self._channel_list[i])
                    self.button_start_list[i].Enable(True)
                else:
                    self.button_start_list[i].Enable(False)

    def port_init_handler(self, arg1):
        self.port_imei_list[arg1["id"]].SetValue("")
        self.port_iccid_list[arg1["id"]].SetValue("")
        self.port_time_list[arg1["id"]].SetValue("")
        self.port_result_list[arg1["id"]].SetValue("")
        self.button_log_list[arg1["id"]].Enable(False)
        # 重复测试结果重置
        self.__test_result[arg1["id"]] = 0
        self.__py_exec_result[arg1["id"]] = ""
        ser = serial_handler.SerialHandler(arg1["PortInfo"])
        imei, iccid = asyncio.run(ser.init_module(arg1["imei"], arg1["iccid"]))
        self.__py_test_fp = open(self.text_ctrl_py.GetValue(), "r", encoding="utf8")
        self.message_queue.put(
            {"id": arg1["id"], "msg_id": "PortTest", "PortInfo": self.port_ctrl_list[arg1["id"]].GetValue(), "Imei": imei, "Iccid": iccid,
             "script": self.__py_test_fp, "time": self.port_time_list[arg1["id"]], "result": self.port_result_list[arg1["id"]]})

    def port_test_handler(self, arg1):
        ser = serial_handler.SerialHandler(arg1["PortInfo"])
        # 测试时间&结果页面刷新  线程ID
        test_res_display = threading.Thread(target=self.__test_time_bar, args=(arg1["id"], arg1["time"], arg1["result"]))
        test_res_display.start()
        ser.write_module(arg1["script"], self.__exec_py_cmd_list)  # 写入脚本开始测试
        ser.ret_result()

        ret_result = []
        log = ""
        length = len(self.testFunctions)
        for i in range(length):
            testFunction = self.testFunctions[i]
            message = self.testMessages[i]

            cmd = "TestBase." + testFunction[0] + "()"
            ser.exec_cmd(cmd)

            test_result = ser.ret_result()  # get recv list
            log += test_result
            try:
                boolean = test_result.split("\r\n")[1:-1]

                if boolean == []:
                    boolean = ["True"]
                if boolean != ["True"] and boolean != ["False"]:
                    boolean = ["False"]
            except Exception as e:
                print(e)
                boolean = ["False"]

            if testFunction[1] == 1:
                dlg  = wx.MessageBox("      当前测试项为:  " + message + "\r\n\r\n      请确认该测试项 【是否通过】", self.port_ctrl_list[arg1["id"]].GetValue(), wx.YES_NO)
                if dlg == wx.YES:
                    boolean = ["True"]
                else:
                    boolean = ["False"]

            ret_result += boolean
            arg1["result"].SetValue("process: "+str(int(i/length*100))+"%")

        test_results = []
        for i in ret_result:
            if i == "True":
                test_results.append("通过")
            elif i == "False":
                test_results.append("不通过")
            else:
                test_results.append(i)

        test_method = []
        for i in self.testFunctions:
            if i[1] == 1:
                test_method.append("人工测试")
            else:
                test_method.append("自动测试")

        self.__py_exec_result[arg1["id"]] = "\r\n".join([i[0]+ ": " + i[1] + "  " + i[2] for i in list(zip(self.testMessages, test_method, test_results))])   # 设置测试结果
        self.logger.write_file(arg1["PortInfo"], log)

        self.__init_excel()
        # TODO excel写入不要列表形式
        if "False" not in ret_result:
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 1})
            self.__excel_write([arg1["PortInfo"], arg1["Imei"], arg1["Iccid"], "Success", str(list(zip(self.testMessages, ret_result)))])
        else:
            self.message_queue.put({"id": arg1["id"], "msg_id": "PortTestEnd", "result": 2})
            self.__excel_write([arg1["PortInfo"], arg1["Imei"], arg1["Iccid"], "Fail", str(list(zip(self.testMessages, ret_result)))])
        self.button_start_list[arg1["id"]].SetLabel(_("开始"))
        self.button_start_list[arg1["id"]].Enable(True)

    def port_test_end_handler(self, arg1):
        self.__test_result[arg1["id"]] = arg1["result"]   # 设置测试结果
        self.button_log_list[arg1["id"]].Enable(True)
        pub.sendMessage('statusBarUpdate', arg1=["The test has been completed", 1])
        self.__port_det(False)

    def __init_excel(self):
        self.__excel_handler = file_handler.ExcelHandler(PROJECT_ABSOLUTE_PATH + "\\Test-Result.xlsx")  # Init Excel
        rows, columns = self.__excel_handler.get_rows_columns()
        if rows == 1 and columns == 1:
            self.__excel_handler.set_cell_value(1, 1, "No.")
            self.__excel_handler.set_cell_value(1, 2, "Com Port")
            self.__excel_handler.set_cell_value(1, 3, "IMEI")
            self.__excel_handler.set_cell_value(1, 4, "ICCID")
            self.__excel_handler.set_cell_value(1, 5, "Test Result")
            self.__excel_handler.set_cell_value(1, 6, "Test Log")

    def __excel_write(self, result):
        if self.__excel_handler:
            rows, columns = self.__excel_handler.get_rows_columns()
            self.__excel_handler.set_cell_value(rows + 1, 1, rows)
            for i, value in enumerate(result):
                self.__excel_handler.set_cell_value(rows + 1, i + 2, value)

    def __test_time_bar(self, test_id, time_bar, result_bar):
        result_bar.SetValue("Testing")
        for i in range(1, 60):
            if self.__test_result[test_id] == 0:
                time_bar.SetValue(str(i) + "S")
                time.sleep(1)
            elif self.__test_result[test_id] == 1:
                result_bar.SetValue("Test Pass")
                break
            elif self.__test_result[test_id] == 2:
                result_bar.SetValue("Test Fail")
                break
        else:
            result_bar.SetValue("Test the timeout")

    @staticmethod
    def __port_det(flag):
        if flag:
            tSerialDet.exit_event()
            pub.sendMessage('statusBarUpdate', arg1=["Port Detect Close", 1])
        else:
            tSerialDet.enter_event()
            pub.sendMessage('statusBarUpdate', arg1=["Port Detect Start", 1])

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
                        self.port_update_handler(message)
                    elif msg_id == "PortInit":
                        self.port_init_handler(message)
                    elif msg_id == "PortTest":
                        self.port_test_handler(message)
                    elif msg_id == "PortTestEnd":
                        self.port_test_end_handler(message)
                    else:
                        pass
            except Exception as e:
                print(e)

    # end of class FactoryFrame
    def close_window(self, event):
        self.Destroy()
        self.__excel_handler.close()
        time.sleep(0.2)
        wx.GetApp().ExitMainLoop()
        wx.Exit()
        process_name = "taskkill /F /IM " + file_name
        p = subprocess.Popen(process_name, shell=True)


class DialogControl(wx.Dialog):
    def __init__(self, parent, title, log_data):
        super(DialogControl, self).__init__(parent, title=title, size=(350, 400))
        self.data = log_data
        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)
        self.panel_sizer = wx.BoxSizer(wx.VERTICAL)
        self.vbox1 = wx.BoxSizer(wx.VERTICAL)
        self.cmd_content = wx.StaticText(self.panel, wx.ID_ANY, _(u"测试日志:"))
        self.vbox1.Add(self.cmd_content, proportion=1, flag=wx.LEFT | wx.EXPAND,)
        self.text_content = wx.TextCtrl(self.panel, wx.ID_ANY, style=wx.TE_MULTILINE)
        self.text_content.SetMinSize((300, 350))
        self.vbox1.Add(self.text_content)
        # set sizer
        self.panel_sizer.Add(self.vbox1, flag=wx.TOP | wx.EXPAND, border=15)
        self.panel.SetSizer(self.panel_sizer)
        self.main_sizer.Add(self.panel,
                            proportion=1,
                            flag=wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND,
                            border=15
                            )
        # Apply sizer and display dialog panel
        self.SetSizerAndFit(self.main_sizer)
        self.Layout()
        self.Center()
        # 设置log显示内容
        self.text_content.SetValue(self.data)


class MyApp(wx.App):
    languageTab = locale.getdefaultlocale()[0]
    print("languageTab: ", languageTab)
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
    # 串口检测线程
    tSerialDet = serial_list.SerialDetection()
    tSerialDet.setDaemon(True)  # set as deamon, stop thread while main frame exit
    tSerialDet.start()

    file_name = os.path.basename(sys.executable)
    app = MyApp()
    app.MainLoop()
