from openpyxl import Workbook, load_workbook

class ExcelParser:
    def get_lead_office(self):
        #book=self.wb
        ws=self.wb["Project pricing - consulting"]
        return ws['C4'].value

    # Initializing with reference to a filename
    def __init__(self,tempfile):
        self.wb = Workbook()
        self.wb = load_workbook(tempfile)


