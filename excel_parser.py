import logging
from openpyxl import Workbook, load_workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string

class ExcelParser:
    # Support method to get the various alternative core sheets
    def get_sheet(self,sheetletter):
        if sheetletter=="A":
            v04name="Project pricing - consulting"
            v06name="A) Project pricing consulting"
            if v04name in wb.get_sheet_names():
                return v04name
            elif v06name in wb.get_sheet_names():
                return v06name
            else:
                return wb.worksheets[1].title
        elif sheetletter == "B":
            pass
        elif sheetletter == "C":
            pass
        else:
            raise ValueError("Sheet names must be A, B, or C")
            return None

    # Returns lead office of the proposal
    def get_lead_office(self):
        ws=self.wb["Project pricing - consulting"]
        return ws['C4'].value

    # Returns region
    def get_region(self):
        ws=self.wb["Project pricing - consulting"]
        return ws['N4'].value

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

    def get_blended_hourly_rate(self):
        ws = self.wb["Project pricing - consulting"]
        for cell in ws['B']:
            if cell.value == "Blended hourly rate:":
                label_cell=cell
        return label_cell.offset(column=1).value

    # Returns a dict of hours estimated by role
    def get_hours_by_role(self):
        ws=self.wb["Project pricing - consulting"]
        for cell in ws['B']:
            if cell.value=="Role":
                first_row=cell.row + 1
            elif cell.value=="SUBTOTAL":
                last_row=cell.row -1

        roles_hours={"Manager": 0, "SPM": 0, "PM": 0, "Cons":0, "Assoc":0}
        for row in ws.iter_rows(min_row=first_row, max_row=last_row, min_col=2, max_col=2):
            for cell in row:
                for key in roles_hours:
                    if key in cell.value:
                        roles_hours[key]+=int(cell.offset(column=8).value)

        return roles_hours

    #Returns total number of hours estiamted
    def get_total_hours(self):
        ws=self.wb["Project pricing - consulting"]
        for cell in ws['13']:
            if cell.value=="Total hours used":
                total_hours_col=column_index_from_string(cell.column)
        for cell in ws['B']:
            if cell.value=="SUBTOTAL":
                total_hours_row=cell.row

        return ws.cell(row=total_hours_row, column=total_hours_col).value

    # Attempts to figure out which pricing method was used by the submitter
    def assess_pricing_method(self):

        ws_B=self.wb["B) Activity-role planning"]
        ws_C=self.wb["C) Activity-role-week planning"]

        #Check total hours in main sheet
        hours_main = self.get_total_hours()

        #Check total hours in sheet B
        for t in tuple(ws_B.rows):
            for cell in t:
                if cell.value=="TOTAL BY ACTIVITY":
                    target_B_col=column_index_from_string(cell.column)
                if cell.value=="TOTAL HOURS BY ROLE":
                    target_B_row=cell.row

        hours_B=ws_B.cell(row=target_B_row, column=target_B_col).value

        #Check total hours in sheet C
        for t in tuple(ws_C.rows):
            for cell in t:
                if cell.value == "TOTAL BY ACTIVITY":
                    target_C_col = column_index_from_string(cell.column)
                if cell.value == "TOTAL HOURS BY WEEK":
                    target_C_row = cell.row

        hours_C = ws_C.cell(row=target_C_row, column=target_C_col).value

        #Check if week-share method has been used
        # TO BE CONSTRUCTED

        #Final check for which method has been used and return string of that result
        if int(hours_B) in range(int(hours_main*0.95), int(hours_main*1.05)):
            return "Method 3 (Activity-Role)"
        elif int(hours_C) in range (int(hours_main*0.95), int(hours_main*1.05)):
            return "Method 4 (Activity-Role-Week)"
        else:
            return "Method 1 or 2"



    # Initializing with reference to a filename
    def __init__(self,tempfile):
        self.wb = Workbook()
        self.wb = load_workbook(tempfile, data_only=True)
        self.filename=tempfile
