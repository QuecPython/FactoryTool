# Copyright (c) Quectel Wireless Solution, Co., Ltd.All Rights Reserved.
#  
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#  
#     http://www.apache.org/licenses/LICENSE-2.0
#  
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

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
                pass
            else:
                return False
        else:
            return False

    @staticmethod
    def det_file_space():
        if uos.statvfs('usr')[3] > 5:
            # return True
            pass
        else:
            return False

    @staticmethod
    def det_file_space2():
        if uos.statvfs('usr')[3] > 5:
            # return True
            pass
        else:
            return False

    @staticmethod
    def det_file_space3():
        if uos.statvfs('usr')[3] > 5:
            # return True
            pass
        else:
            return False
    # ------该区域为测试代码------

