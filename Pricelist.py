import xlrd
import xlwt
import xlutils


class Pricelist:
	def __init__(self, pricelist_filename, preset):
		self.PATH_TO_SOURCES="./pricelists/"
		self.USD_RUB_EXCHANGE_RATE = input("Введите стоимость 1 USD в руб.: ")

		self.pricelist_filename = pricelist_filename
		self.preset = preset
			
		self.parameters = self.load_params()
		self.original_pricelist = self.load_xls()

		self.processed_pricelist = []
		"""
		Обработка
		"""
		if self.preset['processor']:
			processor = getattr(self, self.preset['processor'], None)
			if callable(processor): processor()



	def load_params(self):
		parameters = {
			"startRow": self.preset["startRow"] or int(input("Введите номер начальной строки: ")),
			"cols": self.preset["cols"] or list( input("Введите названия колонок через запятую (латиницей, без пробелов): ").split(","))
		}
		return parameters

	def load_xls(self):
		book = xlrd.open_workbook(f"{self.PATH_TO_SOURCES}{self.pricelist_filename}")
		sheet = book.sheet_by_index(0)
		rows = []

		for idx in range(self.parameters['startRow'], sheet.nrows):
			original_row = sheet.row(idx)
			formatted_row = {}
			for i, column in enumerate(self.parameters['cols']):
				element = original_row[i]
				value = element.value if element.ctype != 2 else int(element.value)
				formatted_row[column] = value
			rows.append(formatted_row)
		return rows
		
	"""
	Методы обработки цен
	"""


	def taimukakeProcessor(self):
		if (input("Использовать стандартные параметры (y/n): ") == "n"):
			deliveryPrice = float(input("Стоимость доставки: "))
			exchangeTax = float(input("Комиссия за обмен %: "))/100
			senderBankFee = float(input("Комиссия банка-отправителя: "))
			recepientBankFee = float(input("Комиссия банка-получателя: "))
			margin = float(input("Маржа %: "))/100
		else:
			deliveryPrice = 48
			exchangeTax = 0.01
			senderBankFee = 80
			recepientBankFee = 40
			margin = 0.09

		params = {"deliveryPrice":deliveryPrice, "exchangeTax":exchangeTax, "senderBankFee":senderBankFee, "recepientBankFee":recepientBankFee, "margin":margin}

		for item in self.original_pricelist:
			self.processed_pricelist.append(self.taimukakeProcessorItem(item, params))


	def taimukakeProcessorItem(self, item, params):
		suppliers = item['price']
		with_margin = suppliers * (1 + params['margin'])
		comissions = (suppliers * params['exchangeTax'])

		fee = params['senderBankFee'] + params['recepientBankFee']
		clients = with_margin + comissions + params['deliveryPrice']

		for_one = clients + fee
		for_five = clients + fee/5
		for_twenty = clients + fee/20

		item["price"] = {
			"1": round(for_one),
			"5": round(for_five),
			"20": round(for_twenty)
		}

		return item
