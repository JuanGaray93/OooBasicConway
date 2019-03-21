REM  *****  BASIC  *****

'''''''''''''''''''''''Declaracion de constantes'''''''''''''''''''''''''
Const COLUMNA_INICIAL = 1
Const FILA_INICIAL = 1
Const ANCHO_TABLERO = 20
Const ALTO_TABLERO = 20
Const COLUMNA_FINAL = COLUMNA_INICIAL + ANCHO_TABLERO - 1
Const FILA_FINAL = FILA_INICIAL + ALTO_TABLERO - 1

Sub Main
	Conway
End Sub

'Pinta el tablero de gris con los valores de cada celda en 1
Function inicializarTablero(hoja)	
	Dim i As Integer
	Dim j as Integer 
	For i = COLUMNA_INICIAL To COLUMNA_FINAL
		For j = FILA_INICIAL To FILA_FINAL
			celda = hoja.getCellByPosition(i, j)
			If celda.Value = 1 Then
				celda.CellBackColor = RGB(0,0,0)
			Else
				If celda.Value = 0 Then
					celda.CellBackColor = RGB(250,250,250)
				Else
					celda.CellBackColor = RGB(200,200,200)
				End If
			End If	
		Next j
	Next i
End Function

'Hace que todas las celdas del mapa sean cero
Function inicializarMapa(mapa)
	Dim i As Integer
	Dim j as Integer 
	For i = 0 To ANCHO_TABLERO - 1
		For j = 0 To ALTO_TABLERO - 1
			mapa(i , j) = 0			
		Next j
	Next i
End Function

'Se le pasa la hoja y las coordenadas, devuelve cuantas celulas vivas vecinas tiene
Function getNumeroDeVecinos(hoja, columna, fila)
	Dim vecinos As Integer
	vecinos = 0
	Dim i
	Dim j
	For i = -1 To 1
		For j = -1 to 1
			If (i <> 0) OR (j <> 0) Then
				celdaRondante = hoja.getCellByPosition(columna + i, fila + j)			
				If( celdaRondante.Value = 1 ) Then
					vecinos = vecinos + 1
				End If
			End If
		Next j
	Next i
	devolverNumeroDeVecinos = vecinos
End Function

Function mapearTablero(hoja, mapa)
	inicializarMapa(mapa)
	Dim circundantes As Integer
	Dim estaVivo As Boolean	
	Dim esCelula As Boolean
	
	Dim estado
	Dim mapeo
	
	Dim iMapa As Integer
	Dim jMapa As Integer
	
	Dim i
	Dim j
	For i = 0 To ANCHO_TABLERO - 1
		For j = 0 to ALTO_TABLERO - 1
			
			iHoja = i + COLUMNA_INICIAL
			jHoja = j + FILA_INICIAL
			estado = hoja.getCellByPosition(iHoja,jHoja).Value
			
			If estado = 1 Then
				estaVivo = True
				esCelula = True
			Else
				If estado = 0 Then
					estaVivo = False
					esCelula = True
				Else
					esCelula = False
				End If
			End If
			
			circundantes = devolverNumeroDeVecinos(hoja, iHoja, jHoja)
			
			If esCelula Then
				If estaVivo Then
					If circundantes < 2 OR circundantes > 3 Then
						mapa(i, j) = -1
					End If
				Else
					If circundantes = 3 Then
						mapa(i, j) = 1
					End If
				End If
			End If
		Next j
	Next i
End Function

Function actualizarTablero(hoja, mapa)
	Dim i
	Dim j
	Dim mapeo
	For i = 0 To ANCHO_TABLERO - 1
		For j = 0 to ALTO_TABLERO - 1
			mapeo = mapa(i,j)
			celda = hoja.getCellByPosition(i + COLUMNA_INICIAL,j + FILA_INICIAL)
			If mapeo = -1 Then
				celda.Value = 0
				celda.CellBackColor = RGB(250,250,250)
			Else
				If mapeo = 1 Then
					celda.Value = 1
					celda.CellBackColor = RGB(0,0,0)
				End If
			End If
		Next j 
	Next i 
End Function

'Juego en si
Sub Conway
	'necesario para poder acceder a las celdas
	esteDocumento = ThisComponent
	hojas = esteDocumento.getSheets()
	hoja = hojas.getByIndex(0)
	'
	inicializarTablero(hoja)
	'mapa es una matriz de int que pueden ser -1, 1 o 0.
	' 0 representa que no hace falta cambiar nada en la celda representada.
	' 1 representa que se debe encender.
	' -1 que se debe apagar
	Dim mapa(ANCHO_TABLERO - 1, ALTO_TABLERO - 1) As Integer
	'Ciclo ppal, corre el juego dentro de esto
	Dim ciclos As Integer
	ciclos = 0
	Do While ciclos < 450
		mapearTablero(hoja, mapa)
		actualizarTablero(hoja, mapa)
		ciclos = ciclos + 1
	Loop
End Sub








