import locale
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Definir a localização para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

driver = webdriver.Chrome()
driver.get('https://www.kabum.com.br/computadores/pc/pc-gamer?gad_source=1&page_number=1&page_size=100&facet_filters=&sort=most_searched')

titulos = driver.find_elements(By.XPATH,'//span[@class="sc-d79c9c3f-0 nlmfp sc-cdc9b13f-16 eHyEuD nameCard"]')
precos = driver.find_elements(By.XPATH,'//span[@class="sc-620f2d27-2 bMHwXA priceCard"]')

workbook = openpyxl.Workbook()
workbook.create_sheet('produtos')
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

# Criar uma lista de tuplas (titulo, preco)
produtos = [(titulo.text, float(preco.text.replace('R$', '').replace('.', '').replace(',', '.'))) for titulo, preco in zip(titulos, precos)]

# Ordenar os produtos pelo preço (segundo elemento da tupla)
produtos_ordenados = sorted(produtos, key=lambda x: x[1])

# Adicionar os produtos ordenados à planilha com formatação de moeda
for i, (titulo, preco) in enumerate(produtos_ordenados, start=2):
    sheet_produtos[f'A{i}'] = titulo
    sheet_produtos[f'B{i}'] = locale.currency(preco, grouping=True)

workbook.save('produtos.xlsx')


