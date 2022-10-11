# 这里写需要导入的模块
import sim
import net
import uos


class TestBase(object):
    def __init__(self):
        self.method_list = [func for func in dir(self) if callable(getattr(self, func)) and not func.startswith("__")]
        self.__run__()

    def __run__(self):
        for i in self.method_list:
            print(eval("TestBase." + i + "()"))

    # ------该区域为测试代码------
    @staticmethod
    def det_signal():
        if sim.getStatus() == 1:
            if net.getConfig()[0] == 5:
                return True
        else:
            return False

    @staticmethod
    def det_file_space():
        if uos.statvfs('usr')[3] > 5:
            return True
        else:
            return False
    # ------该区域为测试代码------


if __name__ == '__main__':
    test_base = TestBase()
