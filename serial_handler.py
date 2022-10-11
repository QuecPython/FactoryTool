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


class SerialPort(object):
    '''
    class for serial operation
    '''
    def __init__(self, port, baud):
        self._port = port
        self._baud = baud
        self._conn = None

    def _open_conn(self):
        try:
            self._conn = serial.Serial(self._port, self._baud)
            return self._conn
        except IOError as e:
            print("Failed reason: %s" % str(e))

    def _close_conn(self):
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

    async def init_module(self, imei_text_ctrl, iccid_text_ctrl):
        self._send_cmd("import sim")
        self._send_cmd("import modem")
        await asyncio.sleep(.1)
        self._conn.flushInput() # 丢弃接收缓存中的所有数据
        try:
            self._send_cmd("modem.getDevImei()")
            await asyncio.sleep(.1)
            imei = self.ret_result().split("\r\n")[1]
            # print("imei: {}".format(imei))
        except Exception as e:
            print("get imei error: {}".format(e))
            imei = "-1"
        finally:
            imei_text_ctrl.SetValue(imei[1:-1])
        try:
            self._send_cmd("sim.getIccid()")
            await asyncio.sleep(.1)
            iccid = self.ret_result().split("\r\n")[1]
            # print("iccid: {}".format(iccid))
            if iccid == "-1":
                iccid = "'No SIM card'"
        except Exception as e:
            print("get iccid error: {}".format(e))
            iccid = "-1"
        finally:
            iccid_text_ctrl.SetValue(iccid[1:-1])
        return imei[1:-1], iccid[1:-1]


    def write_module(self, source, py_cmd, filename="test.py"):
        try:
            if self._test_conn():
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
            else:
                print("串口异常")
        except Exception as e:
            print("文件写入异常：{}".format(e))
        finally:
            source.close()
            return self.exec_py(py_cmd)

    def exec_py(self, cmd):
        self._send_cmd(cmd[0])
        self._conn.flushInput()
        self._send_cmd(cmd[1])
        time.sleep(1)
        return self.ret_result()

    async def run_cmd(self,source:list, log_text_ctrl):
        result_list = ""
        for i in source:
            self._send_cmd(i)
            await asyncio.sleep(.1)
            result_list += self.ret_result()
            log_text_ctrl.SetValue(result_list)
        return result_list

    def ret_result(self):
        return self._conn.read(self._conn.inWaiting()).decode('utf8')

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


