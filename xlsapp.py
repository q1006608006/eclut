import os
import sys
import re

from Common.log import logger
import xlsnest

if not __name__ == '__main__':
    logger.warning('not main app,exit now!!!')
    sys.exit(0)

root = os.path.dirname(__file__)
module_path = root + '/modules'
output = root + '/output'

print('''
                                        欢迎来到德莱联盟
''')

def read_step(msg):
    print(msg)
    val = input('请输入编号：')
    print()
    return val

def check(step, allows, args=None):
    if isinstance(list(allows.keys())[0], int):
        try:
            step = int(step)
        except:
            print('请输入数字!!!')
            return False

    if step in allows:
        func = allows[step]
        if not func:
            print('操作未实现')
            return
        while True:
            state = func(args)
            if state != -8:
                break
    else:
        print('请按照提示输入！')

def quit_app(args: object=...):
    print('程序退出......')
    os._exit(0)

def cancel(*args):
    return 0

def add_mod(*args):
    step = read_step(
        '''请选择模板类型
        1. 普通模板
        2. 关系型模板
        3. 取消
        ''')

    def add_common_mod(*args):
        val = input('请将模板拖入对话框并按回车确认(多个文件用逗号或空格隔开): ')

        fs = re.split(',| ',val)
        for v in fs:
            if v == '':
                del fs[fs.index(v)]
        
        if len(fs) == 0:
            print('请传入文件')
            return -8
        
        print(fs)
        print('导入成功！！！')

    def add_relative_mod(*args):
        val = input('请将模板拖入对话框并按回车确认(不支持多文件): ')
        if val == '':
            print('请传入文件')
            return -8

        print(val)
        print('导入成功')

    check(step,{
        1: add_common_mod,
        2: add_relative_mod,
        3: cancel
    })

def download_excel(*args):
    
    ...

while True:
    step = read_step('''请选择操作:
        1. 导出文档
        2. 导入模板
        3. 退出
        ''')
    check(step, {
        1: download_excel,
        2: add_mod,
        3: quit_app}, None)