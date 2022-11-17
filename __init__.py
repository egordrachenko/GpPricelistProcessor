
"""

"""
import os
from pprint import pprint as pp


PRESETS = {
	"taimukake": {
		"startRow": 4,
		"endRow": 14,
		"series": "A", 
		"model": "B",
		"price": "D",
		"availability": "E"
	}
}

def main():
	clearOutput()
	
	pricelistFilename = letUserChooseFile()
	pickedPreset = letUserChoosePreset()

	pricelistFile = parseXlsFile(pricelistFilename, pickedPreset)
	return 1

def clearOutput():
	os.system("clear")
	os.system("cls")

def letUserChooseFile():
  files = os.listdir("./pricelists")
  print("Выберите файл:")
  for idx, fname in enumerate(files):
  	print(f"{idx}. {fname}")
  
  userInput = input()
  userPick = files[int(userInput)]
  return userPick

def letUserChoosePreset():
	listOfPresets = list(PRESETS.keys())
	print("\nВыберите пресет:")
	for idx,preset in enumerate(listOfPresets):
		print(f"{idx}. {preset}")
	return PRESETS[ listOfPresets[int(input())] ]
	
def parseXlsFile(filename, preset):
	book = xlrd.open_workbook(f"./pricelists/{filename}")
	sheet = book.sheet_by_index(0)
	rows = []
	
	startRow = preset["startRow"] or 0
	for idx in range(startRow, sheet.nrows):
		rows.append(sheet.row(idx))
	pp(rows)

if __name__ == "__main__":
    main()
