#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project:FactoryTool 
@File:log_redirect.py
@Author:rivern.yuan
@Date:2022/10/11 9:47 
"""

import os
import datetime


class RedirectErr:
    """"""

    def __init__(self, obj, logpath):
        """Constructor"""
        if not os.path.exists(logpath + "\\logs\\software\\err\\"):
            os.makedirs(logpath + "\\logs\\software\\err\\")

        file_path = logpath + "\\logs\\software\\err\\" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + "_run_err.log"
        self.filename = open(file_path, "a", encoding='utf-8')

    def write(self, text):
        """"""
        if self.filename.closed:
            pass
        else:
            curr_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.filename.write('[' + str(curr_time) + '] ' + text)
            self.filename.flush()

    def flush(self):
        pass


class RedirectStd:
    """"""

    def __init__(self, obj, logpath):
        """Constructor"""
        if not os.path.exists(logpath + "\\logs\\software\\std\\"):
            os.makedirs(logpath + "\\logs\\software\\std\\")
        file_path = logpath + "\\logs\\software\\std\\" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + "_run_std.log"
        self.filename = open(file_path, "a", encoding='utf-8')

    def write(self, text):
        """"""
        if self.filename.closed:
            pass
        else:
            curr_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if text != "":
                self.filename.write('[' + str(curr_time) + '] ' + text)
                self.filename.flush()

    def flush(self):
        pass


# class StandardOutWrite:
#     def write(self, x):
#         old_std.write(x.replace("\n", " [[%s]]\n" % str(datetime.datetime.now())))
