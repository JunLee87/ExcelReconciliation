#!/usr/bin/env python
# coding=UTF-8
"""
@Author: Jun Lee
@LastEditors: Jun Lee
@contact: lijun87@foxmail.com
@file: Demo1.py
@LastEditTime: 2020-08-26 15:49
@Descripttion:
操作步骤
1、读取	《国内订单-7月n.xlsx》文件中的订单号
2、根据订单号，到 《国际公司.xls》或者《工贸公司.xls》文件中查找 快递单号
3、根据快递单号，到《2020.07月迈粟礼对账单》查找金额
4、填写订单号对应的金额
"""
from util.ReadExcel import ReadExcle
from util.WriteExcel import WriteExcel

if __name__ == '__main__':
    orderIdFilePath='./Resources/国内订单-7月n.xlsx'
    expressDeliveryIdFilePath1='./Resources/ExpressDelivery/国际公司.xls'
    expressDeliveryIdFilePath2='./Resources/ExpressDelivery/工贸公司.xls'
    reconciliationFilePath='./Resources/2020.07月迈粟礼对账单.xlsx'

    # 创建 国内订单-7月 读取对象
    orderIdExcel = ReadExcle(orderIdFilePath, '国内')
    # 获取所有订单
    orderIdMap = orderIdExcel.getCoordinate(1, 2)

    # 国际公司 EXCEL 快递单号
    expressDeliveryExcel_int = ReadExcle(expressDeliveryIdFilePath1, '打印单')
    expressDeliveryIdMap1=expressDeliveryExcel_int.getCoordinate(14,1)

    # 工贸公司 EXCEL 快递单号
    expressDeliveryExcel_dom = ReadExcle(expressDeliveryIdFilePath2, '打印单')
    expressDeliveryIdMap2 = expressDeliveryExcel_dom.getCoordinate(14, 1)

    # 对账单找金额
    reconciliationExcel = ReadExcle(reconciliationFilePath, 'Sheet1')
    reconciliationIdMap = reconciliationExcel.getCoordinate(1, 1)

    # 写入EXCEL
    wExcel = WriteExcel(orderIdFilePath,'国内')

    expressDeliveryId = None
    # 1、读取    《国内订单 - 7月n.xlsx》文件中的订单号
    for oder_key,order in orderIdMap.items():
        # 根据订单号，到 《国际公司.xls》或者《工贸公司.xls》文件中查找快递单号
        expressDelivery_Key = expressDeliveryExcel_dom.matchValue_Regex(expressDeliveryIdMap2, order)
        # 如果内国找不到，就去国际找
        if None==expressDelivery_Key:
            expressDelivery_Key = expressDeliveryExcel_int.matchValue_Regex(expressDeliveryIdMap1, order)
            # 如果国际国内都找不到
            if None == expressDelivery_Key:
                print(f'{order} 国际、工贸都找不到快递单号')
                wExcel.setValue(oder_key[0]+1, 44, '国际、工贸都找不到快递单号')
                wExcel.setColour( 21,oder_key[0] + 1, 'ff0000')
                expressDeliveryId=None
            else:
                # 根据匹配到的订单行，获取对应行的快递单号
                expressDeliveryId = expressDeliveryExcel_int.getSingleValue(expressDelivery_Key[0], expressDelivery_Key[1] - 2)
                # print(f'订单号{order} 国际 快递订单 {expressDeliveryId}')
        else:
            # 根据匹配到的订单行，获取对应行的快递单号
            expressDeliveryId = expressDeliveryExcel_dom.getSingleValue(expressDelivery_Key[0], expressDelivery_Key[1] - 2)
            # print(f'订单号{order} 工贸 快递订单 {expressDeliveryId}')

        # 快递单号不为空才去找金额
        if(None != expressDeliveryId):
            # 3、根据快递单号，到《2020.07月迈粟礼对账单》查找金额
            # 根据快递单号 进行查，返回对应的行号和序号
            reconciliation = reconciliationExcel.findColData(1, 1, expressDeliveryId)

            if(None != reconciliation):
                # 获取匹配订单，所在行的金额
                Amount = reconciliationExcel.getSingleValue(reconciliation[0], reconciliation[1] + 7)
                # print(f'快递单{expressDeliveryId} 金额为 {Amount}')

                # 4、填写订单号对应的金额
                wExcel.setValue(oder_key[0]+1, 21, Amount)


            else:
                print(f'找到订单号，但快递订单 为空')
                wExcel.setValue(oder_key[0]+1, 44, f'找到订单号，但快递订单 为空')
                wExcel.setColour(21,oder_key[0]+1,'cc66ff')

    # 保存
    wExcel.saveExcel(orderIdFilePath)











