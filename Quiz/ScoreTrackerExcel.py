from openpyxl import Workbook

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Student Name"
sheet["B1"] = "Score"
sheet["C1"] = "Remarks"

workbook.save("ScoreTracker.xlsx")
print("Excel successfully made!")

