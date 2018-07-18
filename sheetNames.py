import xlrd

class SheetNames:
    
    def __init__(self, fileName='no_file'):
        self.fileName = fileName
    
    def giveNames(self):
        pointSheetObj = []
        TeamPointWorkbook = xlrd.open_workbook(self.fileName)
        pointSheets = TeamPointWorkbook.sheet_names()
        for i in pointSheets:
            pointSheetObj.append(i)
        return pointSheetObj
        