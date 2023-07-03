from excelManager import ExcelManager
from log import Log
from define import *
from excelManager import ExcelManager


class MainAction:
    __orderPath = ''
    __orderSheet = ''
    __invoicePath = ''
    __invoiceSheet = ''


    def setOrderPath(orderPath):
        MainAction.__orderPath = orderPath

    def setInvoicePath(invoicePath):
        MainAction.__invoicePath = invoicePath

    def setOrderSheet(orderSheet):
        MainAction.__orderSheet = orderSheet

    def setInvoiceSheet(invoiceSheet):
        MainAction.__invoiceSheet = invoiceSheet


    def checkConfig(mode):
        if MainAction.__orderPath == '' or MainAction.__orderSheet == '':
            return False
        
        if mode == EXTRACTION_MODE or mode == MAKE_MODE:
            return True
        
        if MainAction.__invoicePath == '' or MainAction.__invoiceSheet == '':
            return False
        
        if mode == COMBINE_MODE:
            return True


    def runMode(mode):
        if mode == COMBINE_MODE and MainAction.checkConfig(mode):
            MainAction.runCompare()

        elif mode == EXTRACTION_MODE and MainAction.checkConfig(mode):
            MainAction.runBlackList()

        elif mode == MAKE_MODE and MainAction.checkConfig(mode):
            MainAction.runMake()

    def runCompare():
        ex = ExcelManager()
        ex.combineExcels(
            MainAction.__orderPath, 
            MainAction.__orderSheet, 
            MainAction.__invoicePath, 
            MainAction.__invoiceSheet)

    def runBlackList():
        ex = ExcelManager()
        ex.getBlackLisOrder(
            MainAction.__orderPath, 
            MainAction.__orderSheet)

    def runMake():
        ex = ExcelManager()
        ex.makeExcel(
            MainAction.__orderPath,
            MainAction.__orderSheet)