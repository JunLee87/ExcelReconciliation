#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: Config.py
@LastEditTime: 2020-08-27 9:02
@Descripttion:
读取yaml的配置文件
"""
import os

import yaml


class Config:
    def __init__(self,filepath=r'../config.yaml'):
        self.data=None

        # 判断文件是否存在
        if os.path.exists(filepath):
            # 读取yaml文件数据
            with open(filepath, 'r', encoding='utf-8') as yamlFile:
                self.data = yaml.safe_load(yamlFile)
                # print(self.data)
        else:
            raise FileNotFoundError("yaml文件不存在")


if __name__ == '__main__':
    config = Config()
    # 获取路径
    print(config.data['order']['fileName'])
    # 获取sheet名
    print(config.data['order']['sheet']['sheetName'])
    # 列名所在行
    print(config.data['order']['sheet']['column']['columnNameRow'])
    # 列表字段名
    print(config.data['order']['sheet']['column']['columnName'])

    # （从1开始）在第几行开始读取数据
    print(config.data['order']['sheet']['dateSatrtRow'])

    # 正则表达式
    print(config.data['expressDelivery']['expressBillRegex'])


