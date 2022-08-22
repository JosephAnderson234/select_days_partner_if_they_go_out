import openpyxl
#import pandas as pd

#se usa el documento excel
# se definen las listas para ser usadas en los bucles
book = openpyxl.load_workbook("data.xlsx")
sheet = book.active
listaAlumnos = []
listaNumeroDeOrden = []
decisión = []
grupoGeneral = {}|
yes = []
no = []
lunes = []
martes = []
miercoles = []
jueves = []
viernes = []
sabado = []
domingo = []

# se guardan todos los datos de los excel para poder tenerlos por separados en las listas creadas

#Esta duncion establece los booleanos de la desición 
def verificar(text):
	if text == "true":
		return True
	else:
		return False

#se guarda los nombres en la lista listaAlumnos
for i in range(1, 26):
	indice = f'A{i}'
	a1 = sheet[indice].value
	listaAlumnos.append(a1)
#print(listaAlumnos)

# se guarda los número de orden en listaNumeroDeOrden
for j in range(1, 26):
	indice = f'B{j}'
	a1 = int(sheet[indice].value)
	listaNumeroDeOrden.append(a1)
#print(listaNumeroDeOrden)

#se guarda si salen o no en decisión
for k in range(1, 26):
	indice = f'C{k}'
	a1 = verificar(sheet[indice].value)
	decisión.append(a1)
print(decisión)

#se establece la relación del número de orden y la desición
for xd in range(0, 25):
	grupoGeneral[listaNumeroDeOrden[xd]]= decisión[xd]
print(grupoGeneral)

#se separan los numeros de orden en el grupo de los que salen[yes] y los que no salen[no]
for r in range(0, 25):
	if grupoGeneral[r+1] == True:
		yes.append(r+1)
	else:
		no.append(r+1)
#print(yes)
#print(no)
#longitudSi = len(yes)
#longitudNo = len(no)
#print(longitudSi, longitudNo)
#alumnosDiasClases = longitudNo/5
#alumnosDiasFeriados = longitudSi/2
#print(alumnosDiasFeriados, alumnosDiasClases)

def saltarElección(number, parOImpar, semana):
	if parOImpar == True:
		limit = 4
		for d in range(number, 4):
			semana.append(yes[d])
		for e in range(0, 3):
			jueves.append(yes[4 + e])
		for f in range(0, 4):
			viernes.append(yes[7 + f])
		for g in range(0, 3):
			sabado.append(yes[11 + g])
		for h in range(0, 4):
			if h + 15  == len(yes):
				break
			domingo.append(yes[15+h])

#Se itera todos los no a la lista lunes
for a in range(0, 4):
	lunes.append(no[a])
#print(lunes)

#Ahora a la lista de martes
for b in range(0, 3):
	if b + 3 == len(no):
		saltarElección(b, False, martes)
		break
	else:
		martes.append(no[b + 3])
print(martes)
#Se itera los días miercoles 
for c in range(0, 4):
	if c + 7 == len(no):
		saltarElección(c, True, miercoles)
		break
	else:
		miercoles.append(no[c])
print(lunes, martes, miercoles, jueves, viernes, sabado, domingo)
#archivo = pd.read_excel("data.xlsx")
#print(archivo)
print("========================Lunes===================")
for nindex in lunes:
	print(listaAlumnos[nindex-1])

print("========================Martes===================")
for nindex in martes:
	print(listaAlumnos[nindex-1])

print("========================Miercoles===================")
for nindex2 in miercoles:
	print(listaAlumnos[nindex2-1])

print("========================Jueves===================")
for nindex2 in jueves:
	print(listaAlumnos[nindex2-1])

print("========================viernes===================")
for nindex2 in viernes:
	print(listaAlumnos[nindex2-1])

print("========================Sábado===================")
for nindex2 in sabado:
	print(listaAlumnos[nindex2-1])

print("========================Domingo===================")
for nindex2 in domingo:
	print(listaAlumnos[nindex2-1])

