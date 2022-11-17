import xlrd
import xlwt
import xlutils


class Pricelist:
	def __init__(pricelistFilename, preset):
		self.PATH_TO_SOURCES="./pricelists/"
			
		self.original_pricelist = loadXls(pricelistFilename)
		self.parameters = loadParams(preset)
		
	
	def loadXls(filename):
		book = xlrd.open_workbook(f"{self.PATH_TO_SOURCES}{filename}")
		sheet = book.sheet_by_index(0)
		rows = []
	
		startRow = preset["startRow"] or 0
		for idx in range(startRow, sheet.nrows):
			original_row = list(sheet.row(idx).values())
			formatted_row = {
				"series": original_row[0],
				"model": original_row[1],
				"price": int(original_row[2]),
				"availability": originalrow[3]
			}
			rows.append(formatted_row)
		return rows
		
		
	def loadParams(preset):
		
		
