# 这里写需要导入的模块
import sim
import net
import uos


class TestBase(object):
    # ------该区域为测试代码------
    @staticmethod
    def det_signal():
        if sim.getStatus() == 1:
            if net.getConfig()[0] == 5:
                return True
            else:
                return False
        else:
            return False

    @staticmethod
    def det_file_space():
        if uos.statvfs('usr')[3] > 5:
            return True
        else:
            return False
    # ------该区域为测试代码------

