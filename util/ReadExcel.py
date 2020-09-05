#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: readExcel.py
@LastEditTime: 2020-08-26 13:35
@Descripttion:
读取EXCEL
"""
import xlrd
import re


class ReadExcle:
    # 读取EXCEL，返回sheet对应
    def __init__(seft, filePath, sheetName):
        # 读取文EXCEL，获取工作薄对象
        seft.wBook = xlrd.open_workbook(filePath)
        # 切换到指定Sheet
        seft.rSheet = seft.wBook.sheet_by_name(sheetName)

    # 获取列表的有效记录数
    def __getColLen(self, col, starRow):
        return len(self.rSheet.col_values(col, starRow))

    # 查找指定列中关键字(完全匹配，区分大小写)
    # 返回一个元组，匹配到值的（行号，列号），否则为None
    def findColData(self, col, starRow, searchKey):
        mapCoordinate = self.getCoordinate(col, starRow)
        # 遍历集合{（行，列），值}
        for key, value in mapCoordinate.items():
            # 字符串之间判断是否相等
            if searchKey == value:
                return key
        return None

    # 查找指定行中关键字(完全匹配，区分大小写)
    # 返回一个元组，匹配到值的（行号，列号），否则为None
    def findRowData(self, searchKey, row=0, starCol=0) -> tuple:
        col = starCol
        while col < len(self.rSheet.row_values(row)):
            if searchKey == self.getSingleValue(row, col):
                # print(searchKey)
                return (row + 1, col + 1)
            else:
                col += 1

        return None

    # 获取指定列中有效数据集合
    # 元素格式为 （行，列）：单元格值
    def getCoordinate(self, col, starRow):
        # 获取列表的有效记录数
        colLen = self.__getColLen(col, starRow)
        row = starRow
        # 遍历该列中所有值
        i = 0
        # 列表值 字典（key为单元格坐标，value为单元格值）
        mapCoordinate = {}
        while i < colLen:
            # mapCoordinate[(row, col)]=self.rSheet.cell_value(row, col)
            mapCoordinate[(row, col)] = self.getSingleValue(row, col)
            row += 1
            i += 1
        return mapCoordinate

    # 正则表达式获取合同号
    # 找不到值，返回一个空的列表
    def getValue_regex(self, content, expressionList) -> list:
        reultList = []
        for expression in expressionList:
            pattern = re.compile(expression)
            # 通过正则表达式获取值
            reultList = pattern.findall(content)
            # 如果找到结果直接返回，否则下一个规则再判断
            if len(reultList):
                return reultList
        # 如果最后都没匹配到，返回一个空数组
        return reultList

    # 通过正则表达式匹配值
    # 返回一个元组，匹配到值的（行号，列号），否则为None
    def matchValue_Regex(self, valueMap, searchKey, expressionList):
        # key为单元格坐标，value为单元格内容
        for key, value in valueMap.items():
            resultList = self.getValue_regex(value, expressionList)
            # 正则表达式匹配到结果才进行比较
            if len(resultList):
                for i in resultList:
                    if searchKey == i:
                        return key
            # 如果匹配到的结果是一个空列表，不进行比较
            else:
                print('{value} 。正则表达式匹配不到值，无法进行比较')
        # 整列数据都找不到合同，返回空
        return None

    # 获取整个单元格
    # 获取单元格的值，自动去除获到到的值首尾空格
    def getSingleValue(self, row, col):
        value = self.rSheet.cell_value(row, col)
        # 如果值是字符串
        if isinstance(value, str):
            # 去除首尾空格
            value = value.strip()
        # 如果是浮点型
        if isinstance(value, float):
            # 将浮点型转为整型，去掉.0结尾；再转为字符串进行比较
            value = str(int(value))
        return value

    # 获取单元格的值，自动去除获到到的值首尾空格
    # 并根据换行符进行切割，以list型式返回切割后的值
    def getMultipleValue(self, row, col):
        returnList = []
        value = self.rSheet.cell_value(row, col)
        # 如果值是字符串
        if isinstance(value, str):
            # 去除首尾空格
            value = value.strip()

            # 根据 换行符进行切割数据
            returnList = value.split('\n')

        # 如果是浮点型
        if isinstance(value, float):
            # 将浮点型转为整型，去掉.0结尾；再转为字符串进行比较
            value = str(int(value))
            returnList.append(value)
        return returnList

    def cell_value(self, row, col):
        # print(f'cell_value({row}, {col})')
        value=self.rSheet.cell_value(row, col)
        # print(f'+++++++{value}')
        return value



# 测试
if __name__ == '__main__':
    filePath = '../Resources/国内订单-7月n.xlsx'
    sheetName = '国内'
    excel = ReadExcle(filePath, sheetName)
    print(excel.cell_value(297,20))
    # col=1
    # starRow=1
    # colMap=excel.getCoordinate(col, starRow)
    # for key,value in colMap.items():
    #     print(key,value)