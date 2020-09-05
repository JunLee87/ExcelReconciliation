#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: ExcelObj.py
@LastEditTime: 2020-08-27 9:52
@Descripttion:
EXCEL 文件对象
"""
import os

from util.ReadExcel import ReadExcle
from util.WriteExcel import WriteExcel


class ExcelObj(ReadExcle, WriteExcel):
    # class ExcelObj(ReadExcle):
    def __init__(self, filePath, sheetName):
        super().__init__(filePath, sheetName)

        # 用于存在列名及对应的单元格位置 key为列名,value为单元格坐标(行，列)
        self.columnNameMap = {}

        # 文件路径
        self.filePath = filePath

        # sheet名
        self.sheetName = sheetName
        # 列信息
        self.columnMap = {}
        # 字段所在行号
        self.columnNameRow = None

        # 从第几行开始读取数据
        self.dateSatrtRow = None

        # EXCEL表格中最后哪一列，是第几列
        self.lastCol = self.rSheet.ncols + 1

        # 合同号正则表达式规则
        self.contractNoRegex = []

    # 获取EXCEL某一列的所有数据及其对应座标
    def getColData(self, colName):
        # 获取指定列名对应的列
        col = self.columnNameMap[colName][1]
        # print(f' {col-1},{self.dateSatrtRow-1}')
        # 获取指定列中所有数据
        return self.getCoordinate(col - 1, self.dateSatrtRow - 1)

    # 创建写入对象
    def writeExcel(self):
        WriteExcel.__init__(self, self.filePath, self.sheetName)

    # 重写父类的保存文件方法
    def saveExcel(self, ):
        self.wBook.save(self.filePath)

    # 获取字段所在单元格所在列
    def getFieldCol(self, key) -> tuple:
        # 值为（行，列)，行列都从1开始算
        # print(self.columnNameMap[self.columnMap[key]])
        return self.columnNameMap[self.columnMap[key]][1]

    # 累计单元格
    def setGrandTotal(self, row, col, value):
        # 读写需要在同一个对象上操作，所以读取数据要使用WriteExcel对象（openpyxl库）操作
        oldValue = self.wSheet.cell(row, col).value
        # 如果已有值，并且是数字，时行累加并设置新的值
        # 否则会视为没有数据，将当前值设置进去
        if isinstance(oldValue, float):
            self.setValue(row, col, oldValue + value)
        else:
            self.setValue(row, col, value)

    def setGrandTotalDetail(self, row, col, newStr):
        # try:
        #     oldContent = self.self.wSheet.cell(row, col).value
        # except (IndexError):
        #     oldContent = ''
        oldContent = self.wSheet.cell(row, col).value
        #  判断是否是一个空的字符串,如果是就直接设置新值
        #  否则进行拼接后重新录入
        if ('' == oldContent) or (None==oldContent):
            self.setValue(row, col, newStr)
        else:
            self.setValue(row, col, oldContent + '\n' + newStr)

    def getFileName(self):
        return os.path.basename(self.filePath)
