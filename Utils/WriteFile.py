#!/usr/bin/env python
# encoding: utf-8
"""
@contact: 1249294960@qq.com
@software: pengwei
@file: WriteFile.py
@time: 2019/11/14 16:58
@desc:
"""
import os
from ruamel import yaml
from Utils.ConfigRead import *


class YamlWrite(object):

    def Write_Yaml(self, filename, value):
        try:
            if filename in '\\':
                filename.replace('\\', '/')
                if not os.path.exists(filename):
                    os.system(r'type nul>{}'.format(filename))
                    # logger.info('新建文件：%s'%filename)
        finally:
            with open(filename, 'w+', encoding='utf-8') as f:
                yaml.dump(value, f, Dumper=yaml.RoundTripDumper)
                f.close()

    # 追加写入
    def Write_Yaml_A(self, filename, value):
        try:
            if filename in '\\':
                filename.replace('\\', '/')
                if not os.path.exists(filename):
                    os.system(r'type nul>{}'.format(filename))
                    # logger.info('新建文件：%s'%filename)
        finally:
            with open(filename, 'a', encoding='utf-8') as f:
                yaml.dump(value, f, Dumper=yaml.RoundTripDumper)
                f.close()

    # 修改参数
    def Write_Yaml_Updata(self, filename, key, value):
        with open(filename) as f:
            content = yaml.safe_load(f)
            content[key] = value
            f.close()
        with open(filename, 'w+', encoding='utf-8') as f:
            yaml.dump(content, f, Dumper=yaml.RoundTripDumper)
            f.close()


if __name__ == '__main__':
    p = CONFIG_PATH+'Parameter.yaml'
    YamlWrite().Write_Yaml_Updata(p, 'Version', '76')