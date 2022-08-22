import openpyxl
#import pandas as pd

#se usa el documento excel
# se definen las listas para ser usadas en los bucles
book = openpyxl.load_workbook("data.xlsx")
sheet = book.active
listaAlumnos = []
listaNumeroDeOrden = []
decisión = []
grupoGeneral = {}
yes = []
no = []
lunes = []
martes = []
miercoles = []
jueves = []
viernes = []
sabado = []
domingo = []
semanas = [lunes, martes, miercoles, jueves, viernes, sabado, domingo]
elecciones = [yes, no]

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
#print(decisión)

#se establece la relación del número de orden y la desición
for xd in range(0, 25):
	grupoGeneral[listaNumeroDeOrden[xd]]= decisión[xd]
#print(grupoGeneral)

#se separan los numeros de orden en el grupo de los que salen[yes] y los que no salen[no]
for r in range(0, 25):
	if grupoGeneral[r+1] == True:
		yes.append(r+1)
	else:
		no.append(r+1)
#print(yes)
#print(no)
# nlunes = 3
# nmartes = 4
# nmiercoles = 3
# njueves = 4
# nviernes = 3
# nsabado = 4
# ndomingo = 4
# se crea la función que anidara los alumnos a los días correxpondientes

#se define la función para los casos donde no se pueda rellenar bien y su estado
state = True

#esta funcion se ejecuta cada vez que se produce una falta de alumnos en cada día y 
# luego se bloquea el estado, solo se aplica cuando hay más que se quedan
def ecepction(index, dia, longitud, choice):
	if state == True:
		#se obtiene el indice del día problematico
		indexOfDay = semanas.index(dia)
		#si el problema surge al principio de un día
		if index == 0:
			#se completa con ese día con el grupo que salen
			for i in range(0, longitud):
				semanas[indexOfDay].append(choice[i])
			#si el día esta entre el grupo de los que ocupan cuatro se procede a completar tood los
			#días faltantes a este
			if indexOfDay + 1 <= 3:
				#se completa hasta el jueves con 4 alumnos
				for a in range(indexOfDay + 1, 4):
					for b in range(0, 4):
						semanas[a].append(choice[int(b + (longitud*(a-2)))])
				#de aqui hasta el domingo con 3
				for a1 in range(0, 3):
					for b2 in range(0, 3):
						semanas[4 + a1].append(choice[(a1)+(b2 + (16 - len(no)))])
			#Si el día está entre los grupos que necesiten solo 3 alumnos se completa respectivamente
			elif indexOfDay +1 > 3:
				for a in range(indexOfDay + 1, 3):
					for b in range(0, 3):
						semanas[a].append(choice[a+(b + (16 - len(no)))])
		else:
			for i in range(index, longitud):
				dia.append(choice[i - index])
			#si el día siguiente esta entre el grupo de los que ocupan cuatro se
			# procede a completar todos los
			#días faltantes a este
			if indexOfDay + 1 <= 3:
				#se completa hasta el jueves con 4 alumnos
				for a in range(indexOfDay + 1, 4):
					for b in range(0, 4):
						semanas[a].append(choice[int(b + (longitud*(a-2)))])
				#de aqui hasta el domingo con 3
				for a1 in range(0, 3):
					for b2 in range(0, 3):
						semanas[4 + a1].append(choice[(a1)+(b2 + (16 - len(no)))])
			#Si el día está entre los grupos que necesiten solo 3 alumnos se completa respectivamente
			elif indexOfDay +1 > 3:
				for a in range(indexOfDay + 1, 3):
					for b in range(0, 3):
						semanas[a].append(choice[a+(b + (16 - len(no)))])
			# for a in range(indexOfDay + 1 , 7):
			# 	for b in range()
	else:
		pass

#La siguiente eception2 tomarpa esto cuando hay más que se vayan que se queden
state2 = True
def eception2(nRestanteDeAlumnos, day):
	#se crea el estado para que la funcion no se corra 2 veces
	def complete(index, day, indexOfElememt):
		if state2 == True:
			if indexOfElememt == 0:
				for i in range(indexOfElememt, 3):
					day.append(yes[i])
				for a in range(semanas.index(day)+1,7):
					for b in range(0, 3):
						semanas[i].append(yes[j+(i-semanas.index(day))*3])
            
			else:
				indiceDelDia = semanas.index(day) + index
				#print(index, day, indexOfElememt)
				for a in range(indexOfElememt, 3):
					day.append(yes[a-indexOfElememt])
					if a == 2:
						for b in range(indiceDelDia, 7):
							for c in range(0, 3):
								semanas[b].append(yes[a-indexOfElememt + (b - indiceDelDia)*3 + c])
		else:
			pass
	#se itera los alumnos restantes
	for i in range(0, 3):
		for j in range(0,3):
			indice = j+(i*3)+nRestanteDeAlumnos
			#print(indice, j, i)
			#print(len(no))
			if indice >=len(no):
				complete(i, semanas[4+i], j)
				global state2
				state2 = False
				break
			else:
				semanas[semanas.index(day) + i+1].append(no[indice])

#casos de los días 3
def dias3(index, dia, eleccion, acumulación):
	for i in range(index, 3):
		if i + acumulación >= len(yes):
			break
		else:
			dia.append(eleccion[i + acumulación])


#casos de los días 4
#se añade un if para poder 
def dias4(index, dia, eleccion, acumulación):
	for i in range(index, 4):
		if i + acumulación >= len(eleccion):
			ecepction(i, dia, 4, yes)
			global state
			state = False
			return False
			break
		else:
			dia.append(eleccion[i + acumulación])


#comenzamos a iterar		
def viernes_domingo():
	for i in range(0, 3):
		dias3(i, semanas[i + 3], yes, i*3)

def lunes_jueves():
	for i in range(0, 4):
		answer = dias4(0, semanas[i], no, i*4)
		if answer == False:
			break
		else:
			if i == 3:
				if len(no) >= 16:
					eception2(16, jueves)
				else:	
					viernes_domingo()
lunes_jueves()
#print(lunes)

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

