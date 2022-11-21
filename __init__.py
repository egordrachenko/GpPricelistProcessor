
"""

"""
import os
from pprint import pprint as pp
# from Pricelist import TaimukakePricelist


PRESETS = {
	"taimukake": {
		"processor": "TaimukakePricelist",
		"startRow": 3,
		"cols": [
			"series",
			"model",
			'lhr',
			"price",
			"availability"
		]
	}
}


def main():
	pricelist_filename = let_user_choose_file()
	picked_preset = let_user_choose_preset()

	module = __import__("Pricelist")
	class_ = getattr(module, picked_preset['processor'])
	pricelist = class_(pricelist_filename, picked_preset)

	#Pricelist.generatePrivateXls()
	pricelist.generate_public_xls()
	return 1


def let_user_choose_file():
	files = os.listdir("./pricelists")
	print("Выберите файл:")
	for idx, fname in enumerate(files):
		print(f"{idx}. {fname}")
	# user_input = input()
	user_input = 0
	user_pick = files[int(user_input)]
	return user_pick


def let_user_choose_preset():
	list_of_presets = list(PRESETS.keys())
	print("\nВыберите пресет:")
	for idx,preset in enumerate(list_of_presets):
		print(f"{idx}. {preset}")
	# return PRESETS[list_of_presets[int(input())]]
	return PRESETS[list_of_presets[0]]


if __name__ == "__main__":
	main()
