#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: GeneratorExcelOjb.py
@LastEditTime: 2020-08-27 16:23
@Descripttion:
Excel对象生成器
一个文件目录就是一个EXCEL对象

"""

import os

from ExcelObj import ExcelObj
from util.Config import Config


class GeneratorExcelOjb:
    def __init__(self,nodeName,filePath=None):

        # 生成配置文件对象
        # ./config.yaml
        configPath=f'.{os.sep}config.yaml'
        config = Config(configPath)

        # 存放excel对象
        self.excelOjb = None

        # 创建excel对象
        # 根据配置文件中的文件名进行创建
        if 'order'==nodeName:
            # ./Resources/国内订单-7月n.xlsx
            filePath=f'.{os.sep}Resources{os.sep}{config.data[nodeName]["fileName"]}'
            # 创建EXCEL对象
            self.__readExcel(filePath, config.data[nodeName]['sheet']['sheetName'])
        # 根据参数传进来的filePath值进行创建
        else:
            # 文件路径直接读
            self.__readExcel(filePath, config.data[nodeName]['sheet']['sheetName'])
            # 如果是快递单目录，设置合同规则
            if 'expressDelivery'==nodeName:
                self.excelOjb.contractNoRegex=config.data[nodeName]['contractNoRegex']

        # 设置excel对象的列信息
        self.excelOjb.columnMap = config.data[nodeName]['sheet']['column']['columnName']
        # 设置excel对象的 字段所在行号
        self.excelOjb.columnNameRow = config.data[nodeName]['sheet']['column']['columnNameRow']
        # 设置excel对象的 从第几行开始读取数据
        self.excelOjb.dateSatrtRow = config.data[nodeName]['sheet']['dateSatrtRow']

        # 获取EXCEL指定列所在单元格位置
        self.__getFieldCol(self.excelOjb)

    # 读取EXCEL
    def __readExcel(self,filePath,sheetName):
        # 判断文件是否存在
        if os.path.exists(filePath):
            # 创建Excel对象
            self.excelOjb =ExcelObj(filePath, sheetName)
        else:
            raise FileNotFoundError("EXCEL文件不存在")

    #获取字段所在列,返回一个字典，值为{字段名，（行，列）}
    def __getFieldCol(self,excelOjb)->dict:
        for key,value in excelOjb.columnMap.items():
            coordinate=None
            coordinate=excelOjb.findRowData(value,excelOjb.columnNameRow-1)
            # print(coordinate)
            if None==coordinate:
                raise ValueError('没有找到名为 {value} 的列')
            else:
                excelOjb.columnNameMap[value]=coordinate


if __name__ == '__main__':
    # filePath='./Resources/ExpressDelivery/123123.xlsx'
    excel = GeneratorExcelOjb('expressDelivery','./Resources/ExpressDelivery/工贸公司.xls')
    print(excel.excelOjb.filePath)
    print(excel.excelOjb.sheetName)
    print(excel.excelOjb.columnMap)
    print(excel.excelOjb.columnNameRow)
    print(excel.excelOjb.dateSatrtRow)
    print(excel.excelOjb.columnNameMap)

    print(excel.excelOjb.contractNoRegex)
    # # 创建写入对象
    # excel.excelOjb.writeExcel()
    # # 设置单元格值
    # excel.excelOjb.setValue(1,5,333333)
    # # 设置底纹
    # excel.excelOjb.setColour(1,5,'E39191')
    # # 保存文件
    # excel.excelOjb.saveExcel()
