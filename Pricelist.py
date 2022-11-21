import xlrd
import xlwt
import xlutils


class TaimukakePricelist:
	def __init__(self, pricelist_filename, preset):
		self.PATH_TO_SOURCES="./pricelists/"
		self.USD_RUB_EXCHANGE_RATE = 64 #float(input("Введите стоимость 1 USD в руб.: "))

		self.pricelist_filename = pricelist_filename
		self.preset = preset
			
		self.parameters = self.load_params()
		self.merged_cells = []
		self.original_pricelist = self.load_xls()

		self.processed_pricelist = []
		self.xls = xlwt.Workbook()
		"""
		Стили
		"""
		xf = xlwt.easyxf
		self.xf = xf
		self.series_style = xf("font: bold on, height 280; align: vert center, horiz center; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;")
		self.price_style = xf("font: bold on, height 240; align: vert center, horiz right; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;", num_format_str="#,### ₽")
		self.model_style = xf("font: height 240; align: vert center, horiz left; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;")
		self.avail_style = xf("font: height 240; align: vert center, horiz right; borders: top_color black, bottom_color black, right_color black, left_color black, left thin, right thin, top thin, bottom thin;")

		"""
		Обработка
		"""
		self.processor()


	def load_params(self):
		parameters = {
			"startRow": self.preset["startRow"] or int(input("Введите номер начальной строки: ")),
			"cols": self.preset["cols"] or list( input("Введите названия колонок через запятую (латиницей, без пробелов): ").split(","))
		}
		return parameters

	def load_xls(self):
		book = xlrd.open_workbook(f"{self.PATH_TO_SOURCES}{self.pricelist_filename}", formatting_info=True)
		sheet = book.sheet_by_index(0)
		rows = []

		self.merged_cells = self.get_merged_cells(sheet.merged_cells)

		for idx in range(self.parameters['startRow'], sheet.nrows - self.parameters['startRow'] - 1):
			original_row = sheet.row(idx)
			formatted_row = {}
			for i, column in enumerate(self.parameters['cols']):
				element = original_row[i]
				value = element.value if element.ctype != 2 else int(element.value)
				formatted_row[column] = value
			rows.append(formatted_row)
		return rows

	def processor(self):
		if 0: #input("Использовать стандартные параметры (y/n): ") == "n"):
			deliveryPrice = float(input("Стоимость доставки: "))
			exchangeTax = float(input("Комиссия за обмен %: "))/100
			senderBankFee = float(input("Комиссия банка-отправителя: "))
			recepientBankFee = float(input("Комиссия банка-получателя: "))
			margin = float(input("Маржа %: "))/100
		else:
			deliveryPrice = 25
			exchangeTax = 0.03
			senderBankFee = 0
			recepientBankFee = 5
			margin = 0.12

		params = {"deliveryPrice":deliveryPrice, "exchangeTax":exchangeTax, "senderBankFee":senderBankFee, "recepientBankFee":recepientBankFee, "margin":margin}

		for item in self.original_pricelist:
			self.processed_pricelist.append(self.processorItem(item, params))

	def processorItem(self, item, params):
		suppliers = item['price']
		with_margin = suppliers * (1 + params['margin'])
		comissions = (suppliers * params['exchangeTax'])

		fee = params['senderBankFee'] + params['recepientBankFee']
		clients = with_margin + comissions + params['deliveryPrice']

		price = (clients + fee) * self.USD_RUB_EXCHANGE_RATE

		return {
			"series": item["series"],
			"model": "   " + item["model"],
			"availability": "в наличии" if item["availability"] == 'in stock' else "закончились",
			"price": round(price)
		}

	def generate_public_xls(self):
		self.public = self.xls.add_sheet("Public", cell_overwrite_ok=True)

		for row_idx, row in enumerate(self.processed_pricelist):
			self.public.write(row_idx, 0, row['series'], style=self.series_style)
			self.public.write(row_idx, 1, row['model'], style=self.model_style)
			self.public.write(row_idx, 2, row['price'], style=self.price_style)
			self.public.write(row_idx, 3, row['availability'], style=self.avail_style)

		for merged in self.merged_cells:
			self.public.merge(*merged)
			print(merged)

		for i in range(len(self.processed_pricelist)):
			row = self.public.row(i)
			row.height_mismatch = True
			row.height = 256 * 3

		series_col = self.public.col(0)
		model_col = self.public.col(1)
		price_col = self.public.col(2)
		avail_col = self.public.col(3)

		series_col.width = 256 * 20
		model_col.width = 256 * 80
		price_col.width = 256 * 40
		avail_col.width = 256 * 20

		self.xls.save("test.xls")

	def get_merged_cells(self, cells):
		merged_cells = []
		for idx, item in enumerate(cells):
			if idx > 1: #skip first 2
				item = list(item)

				'''
				нужно убрать лишние строки с начала документа
				конечная позиция не включительна, поэтому нунжо вычесть из нее 1
				'''
				y_start = item[0] - self.parameters['startRow']
				y_end = item[1] - self.parameters['startRow'] - 1
				'''
				позицию по х забиваем нулями - объединенные ячейки только в первой колонке 
				'''
				x_start, x_end = 0, 0
				merged_cells.append((y_start, y_end, x_start, x_end))
		return merged_cells
