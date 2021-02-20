import openpyxl
from time import sleep

theFile = openpyxl.load_workbook(f'C:\\Gabarito\Gabarito1.xlsx')
currentSheet = theFile['Planilha1']
linhacoluna = 5
linha = 6
linhainicio = 6
linhafinal = 50
coluna = 6
colunainicio = 1
colunafinal = 75

for c in range(1, 75):
	
	gabarito = currentSheet[5][colunainicio].value
	
	for a in range(6, 51):
		resposta = currentSheet[a][colunainicio].value
		
		if gabarito == 'V':
			
			if resposta == 0:
				currentSheet[a][colunainicio].value = 1
				linhainicio += 1
				if linhainicio >= 50:
					linhainicio = 6
			else:
				currentSheet[a][colunainicio].value = 0
				linhainicio += 1
				if linhainicio >= 50:
					linhainicio = 6
		if gabarito == 'F':
			if resposta == 1:
				currentSheet[a][colunainicio].value = 1
				linhainicio += 1
				if linhainicio >= 50:
					linhainicio = 6
			else:
				currentSheet[a][colunainicio].value = 0
				linhainicio += 1
				if linhainicio >= 50:
					linhainicio = 6
	
	colunainicio += 1

print('fim')
theFile.save(f'C:\\Gabarito\Gabarito1.xlsx')
