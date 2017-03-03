from openpyxl import Workbook, load_workbook

class ExcelParser:
    # Returns lead office of the proposal
    def get_lead_office(self):
        ws=self.wb["Project pricing - consulting"]
        return ws['C4'].value

    # Returns region
    def get_region(self):
        ws=self.wb["Project pricing - consulting"]
        return ws['N4'].value

    def get_hours_by_role(self):
        ws=self.wb["Project pricing - consulting"]

    # Returns total project margin
    def get_margin(self):
        ws=self.wb["Project pricing - consulting"]
        for cell in ws['B']:
            if cell.value=="SUBTOTAL":
                margin_row=cell.row
        return ws.cell(row=margin_row, column=13).value

    # Returns total project fee
    def get_project_fee(self):
        ws=self.wb["Project pricing - consulting"]
        for cell in ws['B']:
            if cell.value=="Project fee:":
                project_fee_row=cell.row
        return ws.cell(row=project_fee_row, column=3).value

    # Returns a list of hours estimated by role
    # NOTE: NOT COMPLETE
    def get_hours_by_role(self):
        ws=self.wb["Project pricing - consulting"]
        for cell in ws['B']:
            if cell.value=="Role":
                first_row=cell.row + 1
            elif cell.value=="SUBTOTAL":
                last_row=cell.row -1

        for row in ws.iter_rows(min_row=first_row, max_row=last_row, min_col=2, max_col=2):
            for cell in row:
                print cell

    # Initializing with reference to a filename
    def __init__(self,tempfile):
        self.wb = Workbook()
        self.wb = load_workbook(tempfile, data_only=True)
        self.filename=tempfile


