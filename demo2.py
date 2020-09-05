#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: DataProcessing.py
@LastEditTime: 2020-08-27 15:52
@Descripttion:
处理数理
"""
import os

from GeneratorExcelOjb import GeneratorExcelOjb


class Demo2:
    def __init__(self):
        # 初始化数据
        # 客户订单
        self.orderObj = GeneratorExcelOjb('order').excelOjb
        # 创建写入象
        self.orderObj.writeExcel()

        # 对账表
        self.reconciliationObj = GeneratorExcelOjb('reconciliation').excelOjb

        # 快递单,存放快递单列表对象
        self.expressDeliveryObjList = []
        self.__loopCreatExcelOjb()

    # 根据快递单目录，循环创建EXCEL对象,并将对象加入到列表中
    def __loopCreatExcelOjb(self):
        filePathList = self.__raversetDirectory()
        for path in filePathList:
            self.expressDeliveryObjList.append(GeneratorExcelOjb('expressDelivery', filePath=path).excelOjb)

    # 循环遍历快递单目录中的文件
    # 返回的列表，存放文件路径
    def __raversetDirectory(self):
        filePathList = []
        # 遍历目录
        # ./Resources/ExpressDelivery/
        dir = f'.{os.sep}Resources{os.sep}ExpressDelivery{os.sep}'

        # 判断目录是否存在
        if os.path.exists(dir):
            # 判断目录是否为空
            if 0 != len(os.listdir(dir)):
                # 获取当前目录下所有内容
                files = os.listdir(dir)
                # 遍历文件夹中所有内容
                for file in files:
                    # 拼接为完整路径
                    filePath = os.path.join(dir, file)
                    # 判断当前路径是否文件
                    if os.path.isfile(filePath):
                        # print(filePath)
                        filePathList.append(filePath)
        return filePathList

    # 处理数据
    def dataProcessing(self):

        # 获取指定列中所有订单
        orderIdMap = self.__getColDataCoordinate(self.orderObj, 'orderId')

        # print(orderIdMap)

        # 1、读取    《国内订单 - 7月n.xlsx》文件中的订单号
        # orderIdMap值为{(行，列)，单元格值}
        for oder_key, order in orderIdMap.items():
            # 存放找到快递对应的EXCEL对象
            expressDeliveryExcel = None
            # 根据订单号，到 《国际公司.xls》或者《工贸公司.xls》文件中查找快递单号
            for expressDeliveryExcelObj in self.expressDeliveryObjList:

                # 获取订单列，所有的数据
                expressDeliveryIdMap = self.__getColDataCoordinate(expressDeliveryExcelObj, 'sellerID')
                # 匹配订单ID
                expressDelivery_Key = expressDeliveryExcelObj.matchValue_Regex(expressDeliveryIdMap, order)
                # 如果找到匹配值，跳出循环不再往下找
                # 否则，在下一个EXCEL对象中查找
                if None != expressDelivery_Key:
                    expressDeliveryExcel = expressDeliveryExcelObj
                    break;

            # 存放匹配到的快递单号
            expressDeliveryId = None
            # 如果快递单目中所有文件都找不到，订单
            if None == expressDeliveryExcel:
                print(f'{order} 所有文件都找不到快递单号')
                # 在订单所有行的最后一列填写果
                # self.orderObj.setValue(oder_key[0] + 1, 44, '所有都找不到快递单号')
                self.orderObj.setValue(oder_key[0] + 1, self.orderObj.lastCol, '所有文件都找不到快递单号')
                # 在金额那列，
                # self.orderObj.setColour(oder_key[0] + 1, 21, 'ff0000')
                self.orderObj.setColour(oder_key[0] + 1, self.orderObj.getFieldCol('freight'), 'ff0000')

            else:
                # 根据匹配到的订单行，获取对应行的快递单号
                expressDeliveryId = expressDeliveryExcel.getValue(expressDelivery_Key[0], expressDelivery_Key[1] - 2)
                print(f'订单号{order}  快递订单 {expressDeliveryId}')

                # 快递单号不为空才去找金额
                if (None != expressDeliveryId):
                    # 3、根据快递单号，到《2020.07月迈粟礼对账单》查找金额
                    # 根据快递单号 进行查，返回对应的行号和列号
                    # reconciliation = self.reconciliationObj.findColData(1, 1, expressDeliveryId)
                    print(self.reconciliationObj.dateSatrtRow-1)
                    print( self.reconciliationObj.getFieldCol('expressDeliveryID')-1)
                    reconciliation = self.reconciliationObj.findColData(self.reconciliationObj.dateSatrtRow-1, self.reconciliationObj.getFieldCol('expressDeliveryID')-1, expressDeliveryId)

                    if (None != reconciliation):
                        # 获取匹配订单，所在行的金额
                        Amount = self.reconciliationObj.getSingleValue(reconciliation[0], reconciliation[1] + 7)
                        # print(f'快递单{expressDeliveryId} 金额为 {Amount}')

                        # 4、填写订单号对应的金额
                        # oder_key[0]是列，oder_key[1]行
                        # self.orderObj.setValue(oder_key[0] + 1, 21, Amount)
                        self.orderObj.setValue(oder_key[0] + 1, self.orderObj.getFieldCol('freight'), Amount)
                    else:
                        print(f'找到订单号，但快递订单 为空')
                        # oder_key[0]是列，oder_key[1]行
                        # self.orderObj.setValue(oder_key[0] + 1, 44, f'找到订单号，但快递订单 为空')
                        self.orderObj.setValue(oder_key[0] + 1, self.orderObj.lastCol, f'找到订单号，但快递订单 为空')
                        # self.orderObj.setColour(oder_key[0] + 1, 21,  'cc66ff')
                        self.orderObj.setColour(oder_key[0] + 1, self.orderObj.getFieldCol('freight'), 'cc66ff')
        # 保存
        self.orderObj.saveExcel()

    # 获取指定列中的所有数据和坐标
    # map的值为 {(行，列)，单元格值}
    def __getColDataCoordinate(self, excelObj, colName) -> dict:
        # 获取
        orderIdCol = excelObj.columnNameMap[excelObj.columnMap[colName]][1]
        orderIdStartRow = excelObj.dateSatrtRow
        # print(f'{orderIdCol-1},{orderIdStartRow-1}')
        return excelObj.getCoordinate(orderIdCol - 1, orderIdStartRow - 1)


if __name__ == '__main__':
    d = Demo2()
    d.dataProcessing()
