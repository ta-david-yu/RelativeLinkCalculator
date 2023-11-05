@ECHO OFF
CHCP 65001
ECHO Run Workbook Relative Link Calculator...\n

RelativeLinkCalculator.exe 2100 Input//2022橡膠工業.xlsx Output//Result_2022.xlsx
RelativeLinkCalculator.exe 1200 Input//2022橡膠工業.xlsx Output//Result_2022.xlsx
RelativeLinkCalculator.exe 1900 Input//2022橡膠工業.xlsx Output//Result_2022.xlsx
RelativeLinkCalculator.exe 2000 Input//2022鋼鐵工業.xlsx Output//Result_2022.xlsx
RelativeLinkCalculator.exe 2800 Input//2022金融業.xlsx Output//Result_2022.xlsx

PAUSE