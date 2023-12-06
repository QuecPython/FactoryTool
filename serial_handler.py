#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project:Factory_test 
@File:serial_handler.py
@Author:rivern.yuan
@Date:2022/9/6 15:03 
"""
import serial
import time
import asyncio
import wx


class SerialPort(object):
    '''
    class for serial operation
    '''
    def __init__(self, port, baud):
        self._port = port
        self._baud = baud
        self._conn = None

    def _open_conn(self):
        self._conn = serial.Serial(self._port, self._baud)
        return self._conn

    def _close_conn(self):
        print("================== close ===================")
        self._conn.close()
        self._conn = None

    def _send_cmd(self, cmd):
        self._conn.write((cmd + '\r\n').encode())

    def _recv_cmd(self):
        time.sleep(.2)
        byte_n = self._conn.inWaiting()
        if byte_n > 0:
            return self._conn.read(byte_n)

    def _test_conn(self):
        self._send_cmd("1")
        # recv: b'1\r\n1\r\n>>> ' 测试正常交互
        test_recv = self._recv_cmd()
        if test_recv:
            if len(test_recv) == 10:
                return True
            else:
                print(test_recv)
                raise SerialError(self._port, "串口持续输出中，请检查运行状态")
        else:
            raise SerialError(self._port, "串口堵塞，请检查串口连通性")


class SerialHandler(SerialPort):
    """write test script & read test result """
    def __init__(self, port, baud=115200):
        super(SerialHandler, self).__init__(port, baud)
        self.init()

    def write_module(self, source, py_cmd, filename="test.py"):
        time.sleep(2)
        self._conn.write(b"\x03")
        self._conn.write(b"\x03")
        self._conn.write(b"\x01")
        self._conn.write(b"\x04")
        self._conn.write(b"\x02")
        # await asyncio.sleep(.1)
        time.sleep(2)
        self._conn.flushInput()

        self._test_conn()

        self._conn.write(b"\x01")
        self._send_cmd("f=open('/usr/" + filename + "','wb')")  # 写入文件
        self._send_cmd("w=f.write")
        while True:
            time.sleep(.1)
            data = source.read(255)
            if not data:
                break
            else:
                self._send_cmd("w(" + repr(data) + ")")
                self._conn.write(b"\x04")
        self._send_cmd("f.close()")
        self._conn.write(b"\x04")
        self._conn.write(b"\x02")

        self._conn.write(b"\x01")
        self._conn.write(b"\x04")
        self._conn.write(b"\x02")

        source.close()
        self.exec_cmd(py_cmd[0])
        return True

    def exec_py(self, cmd):
        self._send_cmd(cmd[0])
        self._conn.flushInput()
        self._send_cmd(cmd[1])
        time.sleep(1)
        return self.ret_result()

    def exec_cmd(self, cmd):
        self._send_cmd(cmd)

    async def run_cmd(self,source:list, log_text_ctrl):
        result_list = ""
        for i in source:
            self._send_cmd(i)
            await asyncio.sleep(.1)
            result_list += self.ret_result()
            log_text_ctrl.SetValue(result_list)
        return result_list

    def getImei(self):
        self._send_cmd("import modem")
        # await asyncio.sleep(.1)
        time.sleep(.1)
        self._conn.flushInput() # 丢弃接收缓存中的所有数据
        try:
            self._send_cmd("modem.getDevImei()")
            # await asyncio.sleep(.1)
            imei = self.ret_result().split("\r\n")[1]
            # print("imei: {}".format(imei))
        except Exception as e:
            print("get imei error: {}".format(e))
            imei = "-1"
        return imei

    def ret_result(self, tips_ctrl=None, tips_info=None):
        data = ""
        for i in range(60):
            if tips_info:
                tips_ctrl.SetValue("")

                font1 = wx.Font(15, wx.MODERN, wx.NORMAL, wx.NORMAL, False, 'Consolas')
                tips_ctrl.SetFont(font1)
                tips_ctrl.SetDefaultStyle(wx.TextAttr(wx.RED))
                tips_ctrl.AppendText(_(u"倒计时: ")+str(60-i))

                font1 = wx.Font(10, wx.MODERN, wx.NORMAL, wx.NORMAL, False, 'Consolas')
                tips_ctrl.SetFont(font1)
                tips_ctrl.SetDefaultStyle(wx.TextAttr(wx.RED))
                tips_ctrl.AppendText("\r\n")
                tips_ctrl.AppendText(tips_info)
            data += self._conn.read(self._conn.inWaiting()).decode("utf-8", errors="ignore")
            if data.endswith(">>> "):
                break
            time.sleep(1)

        if tips_info:
            tips_ctrl.SetValue("")

        if not data.endswith(">>> "):
            self._conn.write(b"\x03")
            time.sleep(0.5)
            data += self._conn.read(self._conn.inWaiting()).decode("utf-8", errors="ignore")
        return data

    def exit_test(self):
        self._close_conn()

    def init(self):
        self._open_conn()


class SerialError(Exception):
    '''
    exception for serial blocking connection
    '''
    def __init__(self, _port, _error):
        self._port = _port
        self._error = _error

    def __str__(self):
        return self._port + " " + self._error


if __name__ == '__main__':
    ser = SerialHandler("COM63")


