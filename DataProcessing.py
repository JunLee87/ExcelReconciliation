#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: DataProcessing.py
@LastEditTime: 2020-08-30 10:21
@Descripttion:
"""
import os
from tkinter import messagebox

from GeneratorExcelOjb import GeneratorExcelOjb


class DataProcessing:
    def __init__(self):
        # 初始化数据
        # 客户订单
        self.orderObj = GeneratorExcelOjb('order').excelOjb
        # 创建写入象
        self.orderObj.writeExcel()

        # 获取最后一列
        self.orderEexcl_lastCol = self.orderObj.lastCol
        # 设置新的列名
        self.orderObj.setValue(self.orderObj.columnNameRow, self.orderEexcl_lastCol, '金额明细')
        # self.orderObj.setValue(self.orderObj.columnNameRow, self.orderEexcl_lastCol+1, '错误日志')

        # 对账表
        self.reconciliationObjList = []
        self.__loopCreatExcelOjb(self.reconciliationObjList, 'reconciliation', 'CourierCompanyStatement')

        # 快递单,存放快递单列表对象
        self.expressDeliveryObjList = []
        self.__loopCreatExcelOjb(self.expressDeliveryObjList, 'expressDelivery', 'ExpressDelivery')

        # 总销售明细
        self.salesDetailsObjList = []
        self.__loopCreatExcelOjb(self.salesDetailsObjList, 'salesDetails', 'SalesDetails')

    # 获取指定列中的所有数据和坐标
    # map的值为 {(行，列)，单元格值}
    def __getColDataCoordinate(self, excelObj, colName) -> dict:
        # 获取
        orderIdCol = excelObj.columnNameMap[excelObj.columnMap[colName]][1]
        orderIdStartRow = excelObj.dateSatrtRow
        # print(f'{orderIdCol-1},{orderIdStartRow-1}')
        return excelObj.getCoordinate(orderIdCol - 1, orderIdStartRow - 1)

    # 处理数据
    def dataProcessing(self):
        # ./config.yaml
        loginPath = f'.{os.sep}Log.txt'
        with open(loginPath,'w', encoding='utf-8') as loginFile:

            # 循环 快递公司对账单 中每个EXCEL
            for reconciliationObj in self.reconciliationObjList:
                # print(f'-----快递公司-----{reconciliationObj.getFileName()}-----------------')

                # 1、根据《2020.07月迈粟礼对账单.xlsx》里面的 运单编号
                # 获 运单编号 列中所有数据
                expressBill_IdMap = self.__getColDataCoordinate(reconciliationObj, 'expressDeliveryID')
                # print(expressBill_IdMap)

                # expressBill_IdMap 值为{(行，列)，单元格值}
                for CourierCompanyOder_key, expressBill_Id in expressBill_IdMap.items():
                    # print(f'--------{CourierCompanyOder_key},{expressBill_Id}')
                    # 存放对 销售明细中的快递单号
                    salesDetail_ExpressBillID = None

                    # 循环 总销售明细目录中所有EXCLE对象
                    for salesDetailsObj in self.salesDetailsObjList:
                        # print(f'-----销售明细-----{salesDetailsObj.getFileName()}-----------------')

                        # 2、根据  《2020.07月迈粟礼对账单.xlsx》运单编号 到 export《总销售明细》运单号，找对应的数据
                        # 将《总销售明细》运单号 ，逐一与 对账单中 运单编号 匹配
                        salesDetails_row = salesDetailsObj.dateSatrtRow - 1
                        salesDetails_loc = salesDetailsObj.getFieldCol('expressDeliveryBill') - 1

                        # 查找（销售）运单号
                        salesDetail_ExpressBillID = salesDetailsObj.findColData(salesDetails_loc, salesDetails_row,
                                                                                expressBill_Id)

                        # 如果找到 运单号 ,取出对应行中的 (快递)订单编号,
                        # 否则往下一个销售明细文件继续查找
                        if salesDetail_ExpressBillID:
                            # export《总销售明细》订单编号
                            salesDetails_OrderIdList = salesDetailsObj.getMultipleValue(salesDetail_ExpressBillID[0],
                                                                                  salesDetails_loc - 1)

                            # 如果有多个订单编号，进行循环查找。如果保有一个就查一次
                            for salesDetails_OrderId in salesDetails_OrderIdList:

                                # 3、根据 订单编号，到 《工贸公司.xls》或《国际公司.xls》中找 匹配的 订单编号，找到对应的订单号后根据卖家备注的合同号
                                # 用于存在 查找到的 订单编号
                                expressDelivery_OrderId = None
                                # 遍历 快递单 目录EXCEL
                                for expressDeliveryObj in self.expressDeliveryObjList:
                                    # print(f'-----快递单-----{expressDeliveryObj.getFileName()}-----------------')

                                    expressDelivery_row = expressDeliveryObj.dateSatrtRow - 1
                                    expressDelivery_loc = expressDeliveryObj.getFieldCol('orderID') - 1

                                    # 匹配 订单编号
                                    expressDelivery_OrderId = expressDeliveryObj.findColData(expressDelivery_loc,
                                                                                             expressDelivery_row,
                                                                                             salesDetails_OrderId)
                                    # 找到 订单编号，如果找不到就下一个文件继续找
                                    if expressDelivery_OrderId:
                                        # print(expressDelivery_OrderId)
                                        # 根据 订单编号 的行，和 “卖家备注” 所在列,获取单元格内容
                                        sellerNotes = expressDeliveryObj.getSingleValue(expressDelivery_OrderId[0],
                                                                                        expressDeliveryObj.getFieldCol(
                                                                                      'sellerID') - 1)
                                        # 根据正式表达式规则，获取  “卖家备注”的合同号
                                        # 返回正则表达式中所有匹配到的值，目前只用第一个值
                                        express_ContractNo = expressDeliveryObj.getValue_regex(sellerNotes,
                                                                                               expressDeliveryObj.contractNoRegex)
                                        # 判断是否找到 合同号
                                        # 如果 找到 合同号 就不再继续往下找
                                        if len(express_ContractNo):
                                            # 只用匹配到的第一个合同号
                                            # print(express_ContractNo[0])
                                            # print(f'======={CourierCompanyOder_key},{reconciliationObj.getFieldCol("amount") - 1}')
                                            # 取对象单EXCEL  运单编号，对应的行，和 “金额”列，取快弹 金额 的值
                                            expressAmount = reconciliationObj.cell_value(CourierCompanyOder_key[0],
                                                                                         reconciliationObj.getFieldCol(
                                                                                             'amount') - 1)
                                            # print(f'--订单--{express_ContractNo[0]}-----金额---{expressAmount}')
                                            # 如果取的值不是浮点型，打型金额有问题
                                            if not isinstance(expressAmount, float):
                                                # print(f'{expressBill_Id} 对应的 金额 为非数值')
                                                loginFile.write(f'{reconciliationObj.getFileName()}中 运单编号 {expressBill_Id} 对应的 金额 为非数值\n')
                                            else:
                                                #   4、根据 合同编号 到 《国内订单》的 订单号中查找对应的行
                                                # 获取指定列中所有数据
                                                orderIdMap = self.__getColDataCoordinate(self.orderObj, 'orderId')

                                                shipping_col = self.orderObj.getFieldCol('Shipping') - 1

                                                # 是否全部都搜索过了
                                                searchAllFlag = True
                                                # orderIdMap值为{(行，列)，单元格值}
                                                for oder_key, order_value in orderIdMap.items():
                                                    if express_ContractNo[0] == order_value:
                                                        # 获取订单对应的金额单元格值
                                                        self.orderObj.setGrandTotal(oder_key[0] + 1, shipping_col + 1,
                                                                                    expressAmount)
                                                        # 设置明细
                                                        self.orderObj.setGrandTotalDetail(oder_key[0] + 1,
                                                                                          self.orderEexcl_lastCol,
                                                                                          f'{reconciliationObj.getFileName()} 运单编号：{expressBill_Id},金额{expressAmount}')

                                                        searchAllFlag = False
                                                        # 不再往下找
                                                        # break

                                                if searchAllFlag:
                                                    # 如果匹配不到 单号，将日志输入到EXCEL中
                                                    # 《工贸公司.xls》或《国际公司.xls》 合同号  在 《国内订单-7月n.xlsx》中找不到
                                                    loginFile.write(f'《{expressDeliveryObj.getFileName()}》 订单：{express_ContractNo[0]} 在 《{self.orderObj.getFileName()}》 找不到\n')
                                        else:
                                            # print(f'{expressDeliveryObj.filePath} 订单编号 {expressDelivery_OrderId} 识别不出对应的合同编号')
                                            loginFile.write(f'《{expressDeliveryObj.getFileName()}》 订单编号： {expressDelivery_OrderId} 识别不出对应的合同编号\n')

                                        # # 《工贸公司.xls》或《国际公司.xls》找到 订单编号 不再继续往下找
                                        # break

                                if None == expressDelivery_OrderId:
                                    # 销售明细文件 订单编号 在 《工贸公司.xls》或《国际公司.xls》 都找不到
                                    loginFile.write(f'《{salesDetailsObj.getFileName()} 》运行号: {expressBill_Id}，对应 订单编号： {salesDetails_OrderId} 在所有快递单文件都找不到 \n')
                                    # 找到就不再继续往下找 销售明细

                            #找值就跳出 总销售明细目录中所有EXCLE 查询循环
                            # break



                    # 如果总销售明细目录中所有EXCLE都找不到
                    if None == salesDetail_ExpressBillID:
                        # 打印查找结果
                        # 快递公司对账单 运输单号 在所有 销售明细文件 都找不到
                        loginFile.write(f'《 {reconciliationObj.getFileName()} 》运单： {expressBill_Id} 在所有销售明细都找不到 \n')

            # 保存
            self.orderObj.saveExcel()

            loginFile.write('------------------数据处理完毕！--------------------')

    # 根据目录名中的EXCEL文件，循环创建EXCEL对象,并将对象加入到列表中
    def __loopCreatExcelOjb(self, excelOjbectList, nodeName, FolderName):
        filePathList = self.__raversetDirectory(FolderName)
        for path in filePathList:
            excelOjbectList.append(GeneratorExcelOjb(nodeName, filePath=path).excelOjb)

    # 循环遍历快递单目录中的文件
    # 返回的列表，存放文件路径
    def __raversetDirectory(self, FolderName):
        filePathList = []
        # 遍历目录
        # ./Resources/ExpressDelivery/
        dir = f'.{os.sep}Resources{os.sep}{FolderName}{os.sep}'

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


if __name__ == '__main__':
    data = DataProcessing()
    data.dataProcessing()

    messagebox.showinfo("提示", "数据处理完成！")
