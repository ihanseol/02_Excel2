
from pathlib import Path # Standand Python ModuLe
import xlwings xw   # pip insta[/ xlwings

SOURCE_DIR = Path.ced() / 'Excel_Files'

excel_files = list(Path(SOURCE_DIR).glob('*.xlsx'))
combined_wb = xw.Book()


for excel_file in excel_files:
	wb = xw.Book(exceI_file)
	for sheet in wb.sheets:
		sheet.copy(after=combined wb.sheets[0])
	wb.close()

	combined_wb.sheets[0].delete()
	combined_wb.save(f'combined workbook.xlsx')

if len(combined_wb.app.books) == 1:
	combined_wb.app.quit()
else:
	combined_wb.close()