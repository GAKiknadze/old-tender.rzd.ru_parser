from bs4 import BeautifulSoup
import requests
import csv
import xlwt


URL = 'https://old-tender.rzd.ru/tender-plan/public/ru?layer_id=6055&STRUCTURE_ID=704&planned_year={}&page6055_4813={}'

FILENAME = 'base_{}'


def get_page(url):
	st = True
	r = requests.get(url)
	soup = BeautifulSoup(r.text, 'html.parser')
	h = soup.find('table', class_='Striped')
	if h:
		st = False
	return h, st


def get_items(table):
	TMP = []
	items = table.find_all_next('tr', class_='tenderplan-row gray')
	for i in items:
		perf = i.find('div', class_='tenderplan-hiddencard')
		per = perf.find_all_next('td')
		pre_info = {
			per[0].text: per[1].text,
			per[2].text: per[3].text,
			per[4].text: per[5].text,
			per[6].text: per[7].text,
			per[8].text: per[9].text,
			per[10].text: per[11].text,
			per[12].text: per[13].text,
			per[14].text: per[15].text,
		}
		for x in i.select('div'):
			x.decompose()
		
		main = i.find_all_next('td')

		num = main[0].text
		pred = main[1].text
		nach = main[2].text
		per_razm = main[3].text
		srok = main[4].text
		sposob = main[5].text
		zak = main[6].text

		info = {
			'Номер закупки в ЕИС': num,
			'Предмет договора': pred,
			'Начальная максимальная цена договора': nach,
			'Дата, период размещения': per_razm,
			'Срок исполнения договора': srok,
			'Способ закупки': sposob,
			'Закупка': zak,
		}

		info.update(pre_info)

		TMP.append(info)

	return TMP


def save_to_csv(filename, sheet, li):
	f = csv.writer(open(filename.format(str(sheet) + '.csv'), "a"))
	for i in li:
		try:
			f.writerow(i.values())
		except:
			pass


def save_to_excel(filename, sheet, li):
	book = xlwt.Workbook(encoding="utf-8")
	sheet = book.add_sheet(str(sheet))
	k = 0
	for i in li:
		p = 0
		for key, value in i.items():
			sheet.write(k, p, value)
			p += 1
		print(str(k) + 'ЗАПИСАНО')
		k += 1
	#book.save('C:/' + filename.format(str(sheet) + '.xls'))
	print('СОХРАНЕНИЕ В ФАЙЛ' + filename)
	book.save(filename)


h = ['base_2017.xls', 'base_2018.xls', 'base_2019.xls', 'base_2020.xls', 'base_2021.xls', 'base_2022.xls']
r = 0
for i in [2017, 2018, 2019, 2020, 2021, 2022]:
	j = 1
	POP = []
	while True:
		text, code = get_page(URL.format(i, j))
		if code:
			break
		print('Год: {} Страница: {}'.format(i, j))
		spisok = get_items(text)
		#save_to_csv(FILENAME, i, spisok)
		POP.extend(spisok)
		j += 1
	save_to_excel(h[r], i, POP)
	r += 1
