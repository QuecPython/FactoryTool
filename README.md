# QuecPython产测工具使用指导手册

## 简介

QuecPython产测工具是专用于QuecPython项目对模块功能测试的工具，方便对出厂前已经烧录正式固件的模块进行相关的功能测试，本文档主要介绍如何使用该工具，在电脑使用Windows7及以上系统都可使用。

适用模块：

所有使用QuecPython开发的模块均可使用

## 功能介绍
1. 支持4台设备同步测试
2. 支持自动测试和人工断点测试
3. 可自定义测试脚本，测试方式（自动测试和人工测试）

## 工具获取

[FactoryTool-Tags](https://github.com/QuecPython/FactoryTool/tags)

github上进入FactoryTool项目最新的tag，可以下载可执行包( FactoryTool.zip ) 和 源代码( Source code(zip/tar.gz) )

## 页面介绍

**菜单栏** 包括编辑菜单和日志菜单

**编辑菜单** 用于编辑测试脚本和Excel日志文件

**日志菜单** 用于查看测试日志文件和工具运行日志文件    

**加载栏** 包括 加载测试文件和配置文件，其中测试文件为要测试的测试文件，配置文件为需要测试function，测试的顺序，测试项的名称，测试的方式(人工测试/自动测试)

**测试区域** 包括 端口、测试项名称、测试的方式、测试结果，工具有四个测试栏位最多可一次性同时测试四个设备，也可以单独测试某个设备。

> **注意：测试的模块需要保证交互口不处于堵塞状态，否则测试将失败**

![1692598668142](./media/1692598668142.jpg)

## 测试脚本

测试脚本编写时需要注意不得改动原有模板代码结构，建议编写完之后先在模块中手动运行测试，确认运行结果是否满足预期，测试脚本模板中需要改动的位置有两个

**导入模块**：导入测试中用到的python库

**测试函数**：自行编写测试函数，函数名可以自定义，但是需要是静态方法且无法传参，函数体内容根据测试需求自定义，函数需要有返回值，必须是bool值

测试脚本示例：

```python
# Write the modules that need to be imported here
import sim
import net
import uos
import utime

class TestBase(object):
    # ------This area is for testing code------
    @staticmethod
    def det_signal():
        utime.sleep(2)
        if sim.getStatus() == 1:
            if net.getConfig()[0] == 5:
                return True
            else:
                return False
        else:
            return False

    @staticmethod
    def det_file_space():
        utime.sleep(2)
        if uos.statvfs('usr')[3] > 5:
            return True
        else:
            return False
    # ------This area is for testing code------
```

## 配置文件
配置文件为json文件，需要配置三维数据，分别为: 测项名称， 测项的方法名称， 测试方式（1为人工测试 0为自动测试）

测试脚本示例：

```json
//             测项名称     function      测试方式: 1为人工测试 0为自动测试
{
    "info": [["Memory test", "det_file_space", 1],
             ["Signal test", "det_signal", 0]]
}
```


## 测试原理

产测工具利用python脚本的形式来测试QuecPython产品的业务功能，通过CDC口来执行测试脚本，测试脚本根据测试需求可以自行调整以达成完成生产测试的功能

1. 通过QuecPython的API接口来获取模组参数和运行状态
2. 通过访问正在运行的业务代码中的对象来获取业务运行情况
3. 通过调用业务代码提供的接口才测试业务功能或者硬件功能
4. 通过QuecPython API接口传输配置文件或者写入产品参数

## 测试步骤

**步骤一**：编辑测试文件和配置文件

编辑测试**module_test.py**代码 和 配置文件**sort_setting.json**，测试代码和配置文件案例如上图所示

**步骤二**：打开工具，选择测试文件和配置文件

打开工具后，点击 **选择** 按钮，分别选择上一步编辑好的测试代码和配置文件，若有检测到串口，则该串口下会显示每一项测试名称以及测试方式

![1692599729648](./media/1692599729648.jpg)

**步骤三**：开始测试

点击 **全部开始** 按钮： 开始测试所有已接入的模块

点击 **开始** 按钮： 开始测试对应的那一个模块

若是人工检测会弹窗提示是否成功，可人工点击 **是/否** 来确认该项结果
![1692599919725](./media/1692599919725.jpg)

**步骤四**：查看测试结果

可在测试结果栏位直接看到测试结果：
1. 通过： 绿色背景色，测试结果栏位标记： 通过
2. 不通过： 红色背景色，测试结果栏位标记： 不通过 
3. 测试中： 黄色背景色，表示该测项正在检测中

![1692599801694](./media/1692599801694.jpg)

**步骤五**：查看测试日志

在工具同级目录下生成**Test-Result.xlsx**文件，包含测试项目内容以及测试结果日志

## 测试结果

每一个模块的的测试结果会写入到excel中导出，可以通过菜单栏的 **编辑** 中的 **编辑Excel文件**菜单打开该excel文件

每一次测试会在excel中生成一条测试记录（无论测试成功或者失败），多次测试会有多条记录，根据测试设备的COM口区分

该文件会保存在工具同级目录下，测试过程中自动追加


## 如何拉取仓库

若有二次开发需求，可直接拉取代码仓库

```shell
git clone --recurse-submodules https://github.com/QuecPython/FactoryTool.git -b interventionable
```

## 环境配置

**安装依赖库**

```shell
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
# 如果上面方式安装失败，可以使用单独安装的形式
pip install wxpython -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install pypubsub -i https://pypi.tuna.tsinghua.edu.cn/simple
```
**运行代码**

```shell
python main.py
```
**编译成可执行程序**

```shell
# 如果上面使用单独安装的形式，需要安装一下pyinstaller库
pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
# 编译成exe程序，输出exe目录在./dist/下
pyinstaller -F -w --win-private-assemblies --icon images/quectel.ico -w ./main.py
```
