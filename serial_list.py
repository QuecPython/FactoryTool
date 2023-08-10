#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project:Factory_test 
@File:serial_list.py
@Author:rivern.yuan
@Date:2022/9/6 15:59 
"""

import time
import threading
import serial.tools.list_ports
from pubsub import pub


class SerialDetection(threading.Thread):
    """
    class for detection port list
    """
    def __init__(self):
        super(SerialDetection, self).__init__()
        self.serPort = serial.tools.list_ports
        self._exit = False

    def exit_event(self):
        self._exit = True

    def enter_event(self):
        self._exit = False

    def get_com_number(self, vid_pid, location):
        port_list = []
        for p in list(self.serPort.comports()):
            if p.vid == int(vid_pid.split(":")[0], 16) and p.pid == int(vid_pid.split(":")[1], 16):
                if location in p.location:
                    port_list.append(p.device)
        return port_list

    def run(self):
        serial_list = []
        while True:
            if self._exit:
                pass
            else:
                if serial_list == [] or serial_list != self.serPort.comports():
                    # serial_port = self.get_com_number("0x2C7C:0x6005", "x.8")
                    # serial_port.extend(self.get_com_number("0x2C7C:0x6005", "x.5"))
                    serial_port = self.get_com_number("0x2C7C:0x0901", "x.8")
                    serial_port.extend(self.get_com_number("0x2C7C:0x6001", "x.5"))
                    serial_port.extend(self.get_com_number("0x2C7C:0x6002", "x.5"))
                    # print("设备列表为{}".format(serial_port))
                    pub.sendMessage('serialUpdate', arg1=serial_port)
            time.sleep(1)


if __name__ == '__main__':
    ser = SerialDetection()
    ser.setDaemon(False)
    ser.start()
